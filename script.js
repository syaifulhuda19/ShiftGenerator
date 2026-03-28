/**
 * SISTEM PENJADWALAN SHIFT v3.2
 * Perubahan utama:
 *  - M shift di-pre-assign: WAJIB 5/hari = 2 T2 + 1 T1E + 2 T1I, min 1 per site
 *  - Generate/regen selalu full-week sehingga continuity & jumping terjaga
 *  - Tabel rekap duty: breakdown total + per tier untuk setiap kategori shift
 *  - Kartu komposisi M per hari (hijau=OK, merah=tidak sesuai target)
 */

// ─── CONSTANTS ────────────────────────────────────────────────────
const MAX_P3_P4_TOTAL = 2;
const MAX_M_PER_WEEK  = 2;   // maks M per orang per minggu
const MAX_M_PER_MONTH = 4;   // maks M per orang per bulan (3–4, pakai 4)

// Target M per hari (FIXED: 5 total)
const M_QUOTA = { "Tier 2": 2, "Tier 1 Email": 1, "Tier 1 Inbound": 2 };
const SHIFT_M_TOTAL = 5;

// M di-assign via pre-selection → M: 0 di pool biasa
const WEIGHTS_WEEKDAYS = { "P1":1, "P2":12, "P3":1, "P4":1, "S":12, "M":0 };
const WEIGHTS_WEEKEND  = { "P1":1, "P2":6,  "P3":0, "P4":0, "S":6,  "M":0 };

const DAY_NAMES   = ["Senin","Selasa","Rabu","Kamis","Jumat","Sabtu","Minggu"];
const MONTH_NAMES = ["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"];
const ALL_SITES   = ["Site 1","Site 2","Site 3","Site 4"];

const templates = {
    weekday: {
        pagi:`Jobdesk MOTD Pagi ({tanggal}) : \nEmail / Backup Escalated : 2 \nTelegram / WA Internal : 2 \nTicket Escalated : 2 \nPIC MyBiz & CP Promo : 2 \nPIC OSS Outage/Maintenance : 2 \nOld Ticket: 1`,
        sore:`Jobdesk MOTD Sore ({tanggal}) : \nEmail / Backup Escalated : 1 \nTelegram / WA Internal : 1 \nTicket Escalated : 1 \nPIC MyBiz/OSS/Maintenance : 4`
    },
    weekend:{
        pagi:`MOTD Pagi Weekend ({tanggal}) : \nEmail : 1 \nTelegram : 2 \nEscalated : 1 \nOSS/MyBiz : 4`,
        sore:`MOTD Sore Weekend ({tanggal}) : \nEmail/Telegram : 2 \nEscalated/OSS : 4`
    }
};

// ─── STATE ────────────────────────────────────────────────────────
let karyawanData    = [];
let schedulesByWeek = {};
let weekLastShifts  = {};
let allRequests     = [];
let currentWeekKey  = null;

// ─── DATE UTILS ───────────────────────────────────────────────────
const toDateKey = d =>
    `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;

function parseKey(key) {
    const [y,m,d] = key.split('-').map(Number);
    return new Date(y, m-1, d);
}
function getMondayOfWeek(date) {
    const off = (date.getDay()+6)%7;
    return new Date(date.getFullYear(), date.getMonth(), date.getDate()-off);
}
function getDayIndex(date) { return (date.getDay()+6)%7; }

function getMonthWeeks(year, month) {
    const last   = new Date(year, month+1, 0);
    let   start  = getMondayOfWeek(new Date(year, month, 1));
    const weeks  = [];
    let   wi     = 1;
    while (start <= last) {
        const end = new Date(start.getFullYear(), start.getMonth(), start.getDate()+6);
        weeks.push({ index: wi, start: new Date(start), end, key: toDateKey(start) });
        start.setDate(start.getDate()+7);
        wi++;
    }
    return weeks;
}

function getPrevWeekLastShift(weekKey, nama) {
    const mon  = parseKey(weekKey);
    const prev = toDateKey(new Date(mon.getFullYear(), mon.getMonth(), mon.getDate()-7));
    return (weekLastShifts[prev] && weekLastShifts[prev][nama]) || "OFF";
}

// Map level → short key
function tk(level) {
    return level === "Tier 2" ? "T2" : level === "Tier 1 Email" ? "T1E" : "T1I";
}

// ─── LOCAL STORAGE ────────────────────────────────────────────────
const LS = {
    schedules:   'shiftapp_schedules_v3',
    requests:    'shiftapp_requests_v3',
    lastShifts:  'shiftapp_lastshifts_v3',
    currentWeek: 'shiftapp_currentweek_v3'
};
function saveState() {
    try {
        localStorage.setItem(LS.schedules,   JSON.stringify(schedulesByWeek));
        localStorage.setItem(LS.requests,    JSON.stringify(allRequests));
        localStorage.setItem(LS.lastShifts,  JSON.stringify(weekLastShifts));
        if (currentWeekKey) localStorage.setItem(LS.currentWeek, currentWeekKey);
    } catch(e) { console.warn("Gagal simpan:", e); }
}
function loadState() {
    try {
        const s = localStorage.getItem(LS.schedules);   if (s) schedulesByWeek = JSON.parse(s);
        const r = localStorage.getItem(LS.requests);    if (r) allRequests     = JSON.parse(r);
        const l = localStorage.getItem(LS.lastShifts);  if (l) weekLastShifts  = JSON.parse(l);
        const w = localStorage.getItem(LS.currentWeek); if (w) currentWeekKey  = w;
    } catch(e) { console.warn("Gagal load:", e); }
}

// ─── INIT ─────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
    loadState();
    fetchKaryawan();

    document.getElementById('btn-generate').addEventListener('click', () => {
        if (currentWeekKey) generateShift(currentWeekKey);
        else alert("Pilih minggu terlebih dahulu!");
    });
    document.getElementById('btn-submit-req').addEventListener('click', handleNewRequest);
    document.getElementById('motd-date').addEventListener('change', updateMOTDText);
    document.getElementById('btn-refresh-ui').addEventListener('click', () => {
        if (currentWeekKey) renderTable(currentWeekKey);
    });
    document.getElementById('btn-reset').addEventListener('click', () => {
        if (!confirm("Reset SEMUA jadwal dan request? Data akan hilang permanen!")) return;
        Object.values(LS).forEach(k => localStorage.removeItem(k));
        location.reload();
    });

    document.getElementById('month-picker').addEventListener('change', onMonthChange);
    document.getElementById('week-select').addEventListener('change', onWeekChange);

    let initYear, initMonth;
    if (currentWeekKey) { const d=parseKey(currentWeekKey); initYear=d.getFullYear(); initMonth=d.getMonth(); }
    else { initYear=2026; initMonth=1; }

    const mp = document.getElementById('month-picker');
    mp.value = `${initYear}-${String(initMonth+1).padStart(2,'0')}`;
    populateWeekSelect(initYear, initMonth, currentWeekKey);
});

async function fetchKaryawan() {
    try {
        const res = await fetch('karyawan.json');
        karyawanData = await res.json();
        const sel = document.getElementById('req-nama');
        sel.innerHTML = '';
        karyawanData.forEach(k => sel.innerHTML += `<option value="${k.nama}">${k.nama}</option>`);
        renderRequestTables();
        if (currentWeekKey) renderTable(currentWeekKey);
    } catch(e) {
        document.getElementById('schedule-body').innerHTML =
            `<tr><td colspan="8" class="text-center p-6 text-red-500 font-bold">
             ⚠️ Gagal memuat data. Gunakan Live Server.</td></tr>`;
    }
}

// ─── WEEK SELECTOR ────────────────────────────────────────────────
function onMonthChange() {
    const val = document.getElementById('month-picker').value;
    if (!val) return;
    const [y,m] = val.split('-').map(Number);
    populateWeekSelect(y, m-1, null);
}
function onWeekChange() {
    currentWeekKey = document.getElementById('week-select').value;
    saveState();
    renderTable(currentWeekKey);
}
function populateWeekSelect(year, month, selectedKey) {
    const weeks = getMonthWeeks(year, month);
    const sel   = document.getElementById('week-select');
    sel.innerHTML = '';
    weeks.forEach(w => {
        const opt = document.createElement('option');
        opt.value = w.key;
        opt.text  = `Minggu ${w.index}: Sen ${w.start.getDate()} ${MONTH_NAMES[w.start.getMonth()]} – Min ${w.end.getDate()} ${MONTH_NAMES[w.end.getMonth()]} ${w.end.getFullYear()}`;
        sel.appendChild(opt);
    });
    const found = selectedKey && weeks.find(w => w.key === selectedKey);
    if (found)             { sel.value = selectedKey; currentWeekKey = selectedKey; }
    else if (weeks.length) { currentWeekKey = weeks[0].key; sel.value = currentWeekKey; saveState(); }
    if (currentWeekKey) renderTable(currentWeekKey);
}

// ─── LOCKED MAP ───────────────────────────────────────────────────
function buildLockedMapForWeek(weekKey) {
    const locked = {};
    allRequests.filter(r => r.status==='accepted').forEach(r => {
        const mon = toDateKey(getMondayOfWeek(parseKey(r.tgl)));
        if (mon !== weekKey) return;
        const di = getDayIndex(parseKey(r.tgl));
        if (!locked[r.nama]) locked[r.nama] = {};
        locked[r.nama][di] = r.shift;
    });
    return locked;
}

// ─── REQUESTS ─────────────────────────────────────────────────────
function handleNewRequest() {
    const nama=document.getElementById('req-nama').value, tgl=document.getElementById('req-tgl').value,
          shift=document.getElementById('req-shift').value, alasan=document.getElementById('req-alasan').value;
    if (!tgl) return alert("Pilih tanggal request!");
    allRequests.push({ id: Date.now(), nama, tgl, shift, alasan, status:'pending' });
    saveState(); renderRequestTables();
    document.getElementById('req-tgl').value='';
    document.getElementById('req-alasan').value='';
}

function processRequest(id, newStatus) {
    const req = allRequests.find(r => r.id===id);
    if (!req) return;
    req.status = newStatus;
    saveState(); renderRequestTables();
    if (newStatus==='accepted') {
        const weekKey = toDateKey(getMondayOfWeek(parseKey(req.tgl)));
        if (schedulesByWeek[weekKey]) generateShift(weekKey);
    }
}

function renderRequestTables() {
    const pb = document.getElementById('pending-req-body');
    const hb = document.getElementById('history-req-body');
    pb.innerHTML = ''; hb.innerHTML = '';
    allRequests.forEach(r => {
        if (r.status==='pending') {
            pb.innerHTML += `<tr class="border-b">
                <td class="p-2">${r.nama}</td><td class="p-2">${r.tgl}</td>
                <td class="p-2 font-bold text-blue-600">${r.shift}</td>
                <td class="p-2 italic text-slate-500">${r.alasan||'-'}</td>
                <td class="p-2 text-center">
                    <button onclick="processRequest(${r.id},'accepted')" class="bg-green-100 text-green-700 px-2 py-1 rounded font-bold mr-1 hover:bg-green-200">Accept</button>
                    <button onclick="processRequest(${r.id},'rejected')" class="bg-red-100 text-red-700 px-2 py-1 rounded font-bold hover:bg-red-200">Reject</button>
                </td></tr>`;
        } else {
            const col = r.status==='accepted' ? 'text-green-600' : 'text-red-600';
            hb.innerHTML += `<tr class="border-b text-slate-500 bg-slate-50">
                <td class="p-2">${r.nama}</td><td class="p-2">${r.tgl}</td>
                <td class="p-2">${r.shift}</td>
                <td class="p-2 font-bold uppercase text-[10px] ${col}">${r.status}</td></tr>`;
        }
    });
}

// ═══════════════════════════════════════════════════════════════════
//  GENERATOR UTAMA
// ═══════════════════════════════════════════════════════════════════

/**
 * Hitung jumlah shift M yang sudah dimiliki seorang karyawan
 * di seluruh minggu dalam bulan yang sama dengan weekKey (tidak termasuk weekKey itu sendiri).
 * Digunakan untuk menjaga batas MAX_M_PER_MONTH.
 */
function getMonthlyMCount(nama, weekKey) {
    const refDate = parseKey(weekKey);
    const refYear = refDate.getFullYear();
    const refMonth = refDate.getMonth();
    let count = 0;
    Object.entries(schedulesByWeek).forEach(([wk, weekly]) => {
        if (wk === weekKey) return;  // Minggu aktif dihitung real-time dari mCount
        const wDate = parseKey(wk);
        // Hitung minggu yang tumpang tindih dengan bulan refMonth
        const wEnd = new Date(wDate.getFullYear(), wDate.getMonth(), wDate.getDate() + 6);
        const inMonth = (
            (wDate.getFullYear() === refYear && wDate.getMonth() === refMonth) ||
            (wEnd.getFullYear() === refYear  && wEnd.getMonth()  === refMonth)
        );
        if (!inMonth) return;
        const emp = weekly.find(e => e.nama === nama);
        if (!emp) return;
        count += emp.shifts.filter(s => s === "M").length;
    });
    return count;
}

/**
 * Pre-pilih siapa yang dapat shift M hari ini.
 * Target: 2 Tier 2 + 1 T1 Email + 2 T1 Inbound = 5 total.
 * Prioritas: cover semua 4 site terlebih dahulu.
 * Aturan: hanya Pria, prevShift≠M, mCount < MAX_M_PER_WEEK, tidak ter-lock ke shift lain.
 * Locked M (accepted request) langsung masuk, mengurangi slot tersisa.
 */
function selectMWorkersForDay(weekly, day, weekKey, lockedMap) {
    const selected = new Set();   // indices
    const byTier   = { "Tier 2":[], "Tier 1 Email":[], "Tier 1 Inbound":[] };
    const covSites = new Set();

    // Helper: Fisher-Yates shuffle (in-place, kembalikan array)
    const shuffle = arr => {
        for (let i = arr.length - 1; i > 0; i--) {
            const j = Math.floor(Math.random() * (i + 1));
            [arr[i], arr[j]] = [arr[j], arr[i]];
        }
        return arr;
    };

    // 1. Kumpulkan locked M (approved request) — berlaku untuk semua gender
    weekly.forEach((k, idx) => {
        const lv = lockedMap[k.nama] && lockedMap[k.nama][day];
        if (lv !== "M") return;
        selected.add(idx);
        if (byTier[k.level]) byTier[k.level].push(idx);
        covSites.add(k.site);
    });

    // 2. Build eligibility + SHUFFLE setiap kali dipanggil (keacakan per-tier)
    const eligible = (tier) => {
        const pool = weekly
            .map((k, idx) => ({ k, idx }))
            .filter(({ k, idx }) => {
                if (selected.has(idx)) return false;
                const lv = lockedMap[k.nama] && lockedMap[k.nama][day];
                if (lv !== undefined) return false;       // Locked ke shift lain
                if (k.gender !== "Pria") return false;
                if (k.level !== tier) return false;
                const prev = day===0 ? getPrevWeekLastShift(weekKey,k.nama) : weekly[idx].shifts[day-1];
                // M→M diizinkan — tidak perlu filter prev==="M"
                if (k.mCount >= MAX_M_PER_WEEK) return false;
                // Cek batas bulanan: mCount minggu ini + bulan lain tidak boleh > MAX_M_PER_MONTH
                if (k.mCount + k._monthlyM >= MAX_M_PER_MONTH) return false;
                // Jika besok ter-lock ke shift non-OFF & non-M, tidak bisa assign M hari ini
                if (day < 6) {
                    const nextLocked = lockedMap[k.nama] && lockedMap[k.nama][day + 1];
                    if (nextLocked && nextLocked !== "OFF" && nextLocked !== "C" && nextLocked !== "M") return false;
                }
                return true;
            });
        return shuffle(pool);   // ← SHUFFLE sebelum dikembalikan
    };

    // 3. Site-aware selection: prioritas site yang belum covered,
    //    dengan urutan site juga di-SHUFFLE agar distribusi merata antar minggu
    const pick = (candidates, needed) => {
        const chosen   = [];
        const pool     = [...candidates];  // Sudah ter-shuffle dari eligible()

        // Shuffle urutan site coverage → cegah site yang sama selalu didahulukan
        const siteOrder = shuffle([...ALL_SITES]);

        // First pass: isi dari site yang belum covered (urutan site random)
        for (const site of siteOrder) {
            if (chosen.length >= needed) break;
            if (covSites.has(site)) continue;
            const ci = pool.findIndex(c => c.k.site === site);
            if (ci !== -1) {
                chosen.push(pool.splice(ci, 1)[0].idx);
                covSites.add(site);
            }
        }

        // Second pass: isi sisa slot dari pool (sudah random)
        while (chosen.length < needed && pool.length) chosen.push(pool.shift().idx);
        return chosen;
    };

    // 4. Isi slot tersisa per tier
    ["Tier 2","Tier 1 Email","Tier 1 Inbound"].forEach(tier => {
        const needed = Math.max(0, M_QUOTA[tier] - byTier[tier].length);
        if (!needed) return;
        pick(eligible(tier), needed).forEach(idx => {
            selected.add(idx);
            byTier[tier].push(idx);
            covSites.add(weekly[idx].site);
        });
    });

    return selected;
}

/**
 * Generate SELURUH MINGGU.
 * - M di-pre-assign terlebih dahulu (5/hari, 2T2+1T1E+2T1I, ≥1/site)
 * - Hari diproses berurutan → continuity & anti-jumping terjaga
 * - Accepted request selalu ter-lock (tidak berubah)
 * - ensureOffContinuity menjaga targetOff=2 tanpa menyentuh locked day
 */
function generateShift(weekKey) {
    const TARGET_OFF = 2;
    const startDate  = parseKey(weekKey);
    const lockedMap  = buildLockedMapForWeek(weekKey);

    let weekly = karyawanData.map(k => ({
        ...k,
        shifts:    Array(7).fill(""),
        offCount:  0,
        mCount:    0,                               // M count minggu ini (real-time)
        _monthlyM: getMonthlyMCount(k.nama, weekKey) // M dari minggu lain di bulan ini
    }));

    for (let day = 0; day < 7; day++) {
        const isWknd  = day >= 5;
        const weights = isWknd ? WEIGHTS_WEEKEND : WEIGHTS_WEEKDAYS;

        // ── A. Pre-assign M ──────────────────────────────────────
        const mSet = selectMWorkersForDay(weekly, day, weekKey, lockedMap);
        mSet.forEach(idx => {
            if (weekly[idx].shifts[day] === "") {  // Belum di-set (cegah double)
                weekly[idx].shifts[day] = "M";
                weekly[idx].mCount++;
            }
        });

        // ── B. Hitung P3/P4 dari locked ──────────────────────────
        let p3p4Daily = 0;
        weekly.forEach(k => {
            const lv = lockedMap[k.nama] && lockedMap[k.nama][day];
            if (lv==="P3"||lv==="P4") p3p4Daily++;
        });

        // ── C. Assign shift sisanya ───────────────────────────────
        const shuffled = [...Array(weekly.length).keys()].sort(() => Math.random()-0.5);

        shuffled.forEach(idx => {
            const k = weekly[idx];
            if (k.shifts[day] !== "") return;  // Sudah di-assign (M/locked)

            const prev   = day===0 ? getPrevWeekLastShift(weekKey,k.nama) : k.shifts[day-1];
            const locked = lockedMap[k.nama] && lockedMap[k.nama][day];

            // a. Apply locked non-M
            if (locked !== undefined) {
                k.shifts[day] = locked;
                if (locked==="OFF"||locked==="C") k.offCount++;
                return;
            }
            // b. Anti-Jumping: setelah M → wajib OFF kecuali sudah dapat M lagi (handled di pre-assign)
            //    Jika sampai sini berarti orang ini tidak dapat M hari ini → paksa OFF
            if (prev === "M") {
                k.shifts[day] = "OFF";
                k.offCount++;
                return;
            }

            // d. Probabilitas OFF natural
            if (k.offCount < TARGET_OFF) {
                if (Math.random() < (isWknd ? 0.6 : 0.1)) {
                    k.shifts[day] = "OFF"; k.offCount++; return;
                }
            }

            // e. Pilih dari pool berbobot — isJumping memfilter semua pelanggaran
            const pool = buildPool(weights, k.gender, prev, p3p4Daily);

            // Fallback bertingkat: cari shift valid yang paling longgar
            let sel;
            if (pool.length) {
                sel = pool[Math.floor(Math.random() * pool.length)];
            } else {
                // Tidak ada shift valid → paksa OFF
                sel = "OFF";
            }
            if (sel === "P3" || sel === "P4") p3p4Daily++;
            if (sel === "OFF") k.offCount++;
            k.shifts[day] = sel;
        });
    }

    // ── D. Koreksi jumlah OFF ─────────────────────────────────────
    ensureOffContinuity(weekly, TARGET_OFF, weekKey, lockedMap);

    // ── E. Simpan tracker Minggu untuk continuity minggu depan ────
    if (!weekLastShifts[weekKey]) weekLastShifts[weekKey] = {};
    weekly.forEach(k => { weekLastShifts[weekKey][k.nama] = k.shifts[6]; });

    schedulesByWeek[weekKey] = weekly;

    const sel = document.getElementById('week-select');
    if (sel.querySelector(`option[value="${weekKey}"]`)) sel.value = weekKey;
    currentWeekKey = weekKey;
    saveState();
    renderTable(weekKey);
}

/**
 * Cek apakah transisi prevShift → nextShift melanggar aturan anti-jumping.
 * Hierarki: M(22) > S(14) > P4(12) > P3(10) > P2(08) > P1(06)
 *   Setelah M  → hanya OFF / C boleh
 *   Setelah S  → tidak boleh P1, P2, P3, P4
 *   Setelah P4 → tidak boleh P1, P2, P3
 *   Setelah P3 → tidak boleh P1, P2
 *   Setelah P2 → tidak boleh P1
 */
function isJumping(prev, next) {
    if (!prev || prev === "OFF" || prev === "C" || prev === "") return false;
    const forbidden = {
        // Setelah M → hanya M / OFF / Cuti boleh (M→M dihandle pre-assign)
        "M":  ["P1","P2","P3","P4","S"],
        // Setelah S → tidak boleh shift P apapun
        "S":  ["P1","P2","P3","P4"],
        // Setelah P4 → tidak boleh P3, P2, P1
        "P4": ["P1","P2","P3"],
        // P3 boleh ke P1/P2 → tidak ada forbidden
        // P2 boleh ke P1   → tidak ada forbidden
    };
    return (forbidden[prev] || []).includes(next);
}

/**
 * Build pool shift berbobot (M selalu dikecualikan).
 * Menggunakan isJumping untuk memfilter semua pelanggaran anti-jumping.
 */
function buildPool(weights, gender, prevShift, p3p4Daily) {
    const pool = [];
    Object.entries(weights).forEach(([s, w]) => {
        if (!w || s === "M") return;
        if (gender === "Wanita" && (s === "P4" || s === "S")) return;
        if (isJumping(prevShift, s)) return;
        if (p3p4Daily >= MAX_P3_P4_TOTAL && (s === "P3" || s === "P4")) return;
        for (let i = 0; i < w; i++) pool.push(s);
    });
    return pool;
}

/**
 * Validasi & koreksi jumlah OFF per minggu.
 * - Locked day TIDAK boleh diubah sama sekali.
 * - OFF wajib setelah M tidak boleh diubah.
 */
function ensureOffContinuity(weekly, targetOff, weekKey, lockedMap) {
    weekly.forEach(k => {
        // Hitung ulang offCount setelah semua shift terisi
        k.offCount = k.shifts.filter(s => s==="OFF"||s==="C").length;
        const emp  = lockedMap[k.nama] || {};

        // Terlalu banyak OFF → ubah ke shift valid (respek anti-jumping)
        while (k.offCount > targetOff) {
            let changed = false;
            for (let i = 0; i < 7; i++) {
                if (k.shifts[i] !== "OFF") continue;
                if (emp[i] !== undefined)  continue;                 // Locked → skip
                const prev = i === 0 ? getPrevWeekLastShift(weekKey, k.nama) : k.shifts[i-1];
                if (prev === "M")          continue;                 // Wajib post-M → skip
                // Cari shift pengganti yang tidak melanggar anti-jumping ke prev & next
                const next = k.shifts[i + 1] || null;
                // Kandidat pengganti: P2 atau S (bobot tinggi), asal tidak jumping
                const candidates = ["P2","S","P1","P3","P4"].filter(s => {
                    if (isJumping(prev, s)) return false;
                    if (k.gender === "Wanita" && (s === "P4" || s === "S")) return false;
                    if (next && isJumping(s, next)) return false;    // Jaga hari berikutnya
                    return true;
                });
                if (!candidates.length) continue;
                k.shifts[i] = candidates[0];
                k.offCount--;
                changed = true;
                break;
            }
            if (!changed) break;
        }

        // Terlalu sedikit OFF → paksa OFF di hari non-locked, non-M, non-C
        while (k.offCount < targetOff) {
            const cands = [];
            k.shifts.forEach((s, i) => {
                if (emp[i] !== undefined)            return;   // Locked → skip
                if (s === "OFF" || s === "M" || s === "C") return;
                if (k.shifts[i+1] === "OFF")         return;   // Hindari double OFF berurutan
                // Setelah OFF, hari berikutnya boleh shift apapun (OFF tidak trigger jumping)
                // Namun pastikan hari ini bukan setelah M (tidak mungkin, M sudah dapat OFF di step atas)
                const prev = i === 0 ? getPrevWeekLastShift(weekKey, k.nama) : k.shifts[i-1];
                if (prev === "M") return;                       // Hari wajib OFF post-M, skip
                cands.push(i);
            });
            if (!cands.length) break;
            const ri = cands[Math.floor(Math.random()*cands.length)];
            k.shifts[ri] = "OFF"; k.offCount++;
        }
    });
}

// ═══════════════════════════════════════════════════════════════════
//  RENDER
// ═══════════════════════════════════════════════════════════════════

function renderTable(weekKey) {
    if (!weekKey) return;

    const tbody   = document.getElementById("schedule-body");
    const thead   = document.getElementById("table-header");
    const motdSel = document.getElementById("motd-date");
    tbody.innerHTML = ''; motdSel.innerHTML = '';

    const startDate = parseKey(weekKey);

    // Markers accepted request untuk minggu ini
    const reqMarkers = new Set();
    allRequests.filter(r => r.status==='accepted').forEach(r => {
        const mon = toDateKey(getMondayOfWeek(parseKey(r.tgl)));
        if (mon !== weekKey) return;
        reqMarkers.add(`${r.nama}||${getDayIndex(parseKey(r.tgl))}`);
    });

    // Header
    let hdr = `<tr><th class="p-2 border border-slate-600 bg-slate-900 w-64 text-left sticky left-0 z-10">Nama</th>`;
    for (let i=0; i<7; i++) {
        const d    = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate()+i);
        const dStr = `${d.getDate()}/${d.getMonth()+1}/${d.getFullYear()}`;
        const wknd = i>=5;
        hdr += `<th class="${wknd?'bg-red-700':'bg-slate-900'} p-2 border border-slate-600">
            ${DAY_NAMES[i].substring(0,3)}<br><span class="text-[10px] font-normal">${dStr}</span></th>`;
        const opt = document.createElement("option");
        opt.value = wknd?"weekend":"weekday";
        const full = `${DAY_NAMES[i]}, ${dStr}`;
        opt.text=full; opt.dataset.fulldate=full;
        motdSel.appendChild(opt);
    }
    thead.innerHTML = hdr+'</tr>';

    // Rows
    const data = schedulesByWeek[weekKey]||[];
    if (!data.length) {
        tbody.innerHTML = `<tr><td colspan="8" class="text-center p-10 text-slate-400 italic text-sm">
            📭 Belum ada jadwal. Klik <strong class="text-blue-600">🎲 Generate Baru</strong>.</td></tr>`;
    } else {
        data.forEach(k => {
            const lvl = k.level==="Tier 2"?"bg-t2":k.level==="Tier 1 Email"?"bg-t1e":"bg-t1i";
            let row = `<tr class="border-b border-slate-300">
                <td class="p-1 border border-slate-400 text-left text-[11px] font-bold uppercase ${lvl} sticky left-0">${k.nama}</td>`;
            k.shifts.forEach((s,i) => {
                const req   = reqMarkers.has(`${k.nama}||${i}`);
                const badge = req ? `<span class="req-badge" title="Approved request">★</span>` : '';
                row += `<td class="p-0 border border-slate-400 text-[11px] font-bold shift-${s} relative${req?' cell-requested':''}">
                    <div class="cell-inner">${badge}${s}</div></td>`;
            });
            tbody.innerHTML += row+'</tr>';
        });
    }

    renderDutySummary(weekKey, startDate);
    updateMOTDText();
}

/**
 * Tabel rekap duty + breakdown per tier.
 * Bonus: kartu komposisi M per hari.
 */
function renderDutySummary(weekKey, startDate) {
    const container = document.getElementById('duty-summary');
    if (!container) return;
    const data = schedulesByWeek[weekKey]||[];
    if (!data.length) { container.innerHTML=''; return; }

    // Hitung per hari, per kategori, per tier
    const cnt = Array.from({length:7}, () => ({
        P:  {total:0,T2:0,T1E:0,T1I:0},
        S:  {total:0,T2:0,T1E:0,T1I:0},
        M:  {total:0,T2:0,T1E:0,T1I:0},
        OFF:{total:0,T2:0,T1E:0,T1I:0},
    }));
    data.forEach(k => {
        const t = tk(k.level);
        k.shifts.forEach((s,i) => {
            const cat = s.startsWith("P") ? "P" : s==="S" ? "S" : s==="M" ? "M"
                      : (s==="OFF"||s==="C") ? "OFF" : null;
            if (!cat) return;
            cnt[i][cat].total++;
            cnt[i][cat][t]++;
        });
    });

    // Header hari
    let colHdr = '';
    for (let i=0; i<7; i++) {
        const d    = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate()+i);
        const dStr = `${d.getDate()}/${d.getMonth()+1}`;
        const wknd = i>=5;
        colHdr += `<th class="${wknd?'text-red-600 bg-red-50':'text-slate-700'} text-center p-2 font-bold text-xs border border-slate-200">
            ${DAY_NAMES[i].substring(0,3)}<br><span class="font-normal text-[10px]">${dStr}</span></th>`;
    }

    const CATS = [
        {key:'P',   label:'☀️ Pagi (P1–P4)', bg:'bg-amber-50',   txt:'text-amber-800'},
        {key:'S',   label:'🌆 Sore (S)',      bg:'bg-blue-50',    txt:'text-blue-800'},
        {key:'M',   label:'🌙 Malam (M)',     bg:'bg-purple-50',  txt:'text-purple-800'},
        {key:'OFF', label:'🏠 OFF / Cuti',    bg:'bg-slate-50',   txt:'text-slate-500'},
    ];
    const TIERS = [
        {t:'T2',  lbl:'↳ Tier 2',     cls:'text-yellow-700'},
        {t:'T1E', lbl:'↳ T1 Email',   cls:'text-red-600'},
        {t:'T1I', lbl:'↳ T1 Inbound', cls:'text-violet-700'},
    ];
    // M targets per tier-key
    const M_TARGET = { T2:2, T1E:1, T1I:2 };

    let rows = '';

    CATS.forEach(cat => {
        // Baris total kategori
        let cells = '';
        cnt.forEach((c,i) => {
            const v    = c[cat.key].total;
            const wknd = i>=5;
            const warn = (cat.key==='M' && v!==SHIFT_M_TOTAL) ? ' text-red-500' : '';
            cells += `<td class="p-2 text-center text-sm font-bold border border-slate-200 ${cat.txt}${warn}${wknd?' bg-red-50/40':''}">${v}</td>`;
        });
        rows += `<tr class="${cat.bg} border-t-2 border-slate-200">
            <td class="p-2 border border-slate-200 text-xs font-extrabold whitespace-nowrap">${cat.label}</td>
            ${cells}</tr>`;

        // Sub-baris per tier
        TIERS.forEach(tier => {
            let sub = '';
            cnt.forEach((c,i) => {
                const v    = c[cat.key][tier.t];
                const wknd = i>=5;
                // Untuk M: highlight jika tidak sesuai target
                const mWarn = (cat.key==='M' && v !== M_TARGET[tier.t]) ? ' text-red-500 font-bold' : '';
                sub += `<td class="p-1.5 text-center text-[11px] border border-slate-200${mWarn}${wknd?' bg-red-50/40':''}">${v}</td>`;
            });
            rows += `<tr class="${cat.bg}">
                <td class="pl-5 p-1.5 border border-slate-200 text-[11px] ${tier.cls} whitespace-nowrap">${tier.lbl}</td>
                ${sub}</tr>`;
        });
    });

    // Total Duty
    let totCells = '';
    cnt.forEach((c,i) => {
        const tot  = c.P.total+c.S.total+c.M.total;
        const wknd = i>=5;
        totCells += `<td class="p-2 text-center text-sm font-extrabold border border-slate-200 bg-slate-100${wknd?' border-red-200':''}">${tot}</td>`;
    });
    rows += `<tr class="border-t-2 border-slate-400 bg-slate-100">
        <td class="p-2 border border-slate-200 text-xs font-extrabold text-slate-800">📊 Total Duty</td>
        ${totCells}</tr>`;

    // Kartu komposisi M per hari
    let mCards = '';
    for (let i=0; i<7; i++) {
        const m  = cnt[i].M;
        const ok = m.T2===2 && m.T1E===1 && m.T1I===2;
        mCards += `
            <div class="flex-1 min-w-[72px] text-center p-2 rounded-lg border ${ok?'border-green-300 bg-green-50':'border-red-300 bg-red-50'}">
                <div class="text-[10px] font-bold text-slate-500 mb-1">${DAY_NAMES[i].substring(0,3)}</div>
                <div class="text-[11px] leading-snug space-y-px">
                    <div class="${m.T2===2  ?'text-green-700':'text-red-500'} font-bold">T2: ${m.T2}</div>
                    <div class="${m.T1E===1 ?'text-green-700':'text-red-500'} font-bold">T1E: ${m.T1E}</div>
                    <div class="${m.T1I===2 ?'text-green-700':'text-red-500'} font-bold">T1I: ${m.T1I}</div>
                </div>
                <div class="text-[10px] mt-1 font-extrabold ${ok?'text-green-600':'text-red-500'}">${ok?'✓':'⚠'} ${m.T2+m.T1E+m.T1I}/5</div>
            </div>`;
    }

    container.innerHTML = `
    <div class="space-y-4">

        <!-- Tabel Rekap Duty -->
        <div class="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
            <div class="p-4 border-b bg-slate-50">
                <h2 class="font-bold text-slate-800">📊 Rekap Agent Duty per Hari</h2>
                <p class="text-xs text-slate-400 mt-0.5">
                    Breakdown per shift & tier.
                    <span class="text-red-500 font-bold">Merah</span> = tidak sesuai target &nbsp;|&nbsp;
                    Target M: <span class="font-bold text-yellow-700">2 T2</span> ·
                               <span class="font-bold text-red-600">1 T1E</span> ·
                               <span class="font-bold text-violet-700">2 T1I</span>
                </p>
            </div>
            <div class="overflow-x-auto p-2">
                <table class="w-full border-collapse text-sm">
                    <thead>
                        <tr>
                            <th class="text-left p-2 text-xs font-bold text-slate-500 border border-slate-200 bg-slate-50 w-36">Shift / Tier</th>
                            ${colHdr}
                        </tr>
                    </thead>
                    <tbody>${rows}</tbody>
                </table>
            </div>
        </div>

        <!-- Kartu Komposisi M -->
        <div class="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
            <div class="p-4 border-b bg-slate-50">
                <h2 class="font-bold text-slate-800">🌙 Komposisi Shift Malam per Hari</h2>
                <p class="text-xs text-slate-400 mt-0.5">
                    Target harian:
                    <span class="font-bold text-yellow-700">2 Tier 2</span> ·
                    <span class="font-bold text-red-600">1 T1 Email</span> ·
                    <span class="font-bold text-violet-700">2 T1 Inbound</span> ·
                    min. 1 per site
                </p>
            </div>
            <div class="p-4 flex flex-wrap gap-2">${mCards}</div>
        </div>

    </div>`;
}

function updateMOTDText() {
    const sel = document.getElementById("motd-date");
    if (!sel||!sel.options[sel.selectedIndex]) return;
    const type = sel.value, date = sel.options[sel.selectedIndex].dataset.fulldate;
    document.getElementById("motd-pagi-text").value = templates[type].pagi.replace("{tanggal}",date);
    document.getElementById("motd-sore-text").value = templates[type].sore.replace("{tanggal}",date);
}
