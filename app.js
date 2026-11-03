/* Dashboard Keputusan Pin Akrilik - Praktikum Terintegrasi II */
const $ = (s, r = document) => r.querySelector(s);
const fmt = (n, d = 0) => Number(n).toLocaleString('id-ID', { minimumFractionDigits: d, maximumFractionDigits: d });
const L = ['Tidak aman', 'Sedang', 'Aman'], K = ['bad', 'mid', 'ok'], COL = ['#d64545', '#e0a100', '#1f9d55'];
const reduce = matchMedia('(prefers-reduced-motion: reduce)').matches;

// Parameter awal: semuanya dari modul, kecuali yang ditandai (asumsi)
const P = {
  std: 8000, cus: 10000, vStd: 2313.21, vCus: 3813.21,      // Modul II
  fixed: 777500, order: 15000, angsuran: 128831,           // Modul II & III
  target: 200, cacat: 5, warnCacat: 4,                      // SOP-05, SOP-08, warnCacat = asumsi
  ss: 51, rop: 101,                                         // Modul III
  w: [30, 30, 20, 20],                                      // Modul IV (C1, C2, C3, C4)
  aman: 0.75, sedang: 0.6, budget: ''                       // asumsi
};
const PARAMS = [
  ['target', 'Target normal (pcs/bulan)'], ['fixed', 'Biaya tetap (Rp/bulan)'], ['order', 'Biaya pemesanan (Rp/bulan)'],
  ['angsuran', 'Angsuran bank (Rp/bulan)'], ['cacat', 'Batas cacat (%)'], ['warnCacat', 'Mulai waspada cacat (%)'],
  ['ss', 'Safety stock (pin)'], ['rop', 'Reorder point (pin)'],
  ['w0', 'Bobot C1 Waktu penyelesaian (%)'], ['w1', 'Bobot C2 Throughput (%)'], ['w2', 'Bobot C3 Waktu tunggu (%)'], ['w3', 'Bobot C4 Biaya tambahan (%)'],
  ['aman', 'Skor V dianggap Aman ≥'], ['sedang', 'Skor V dianggap Sedang ≥'], ['budget', 'Batas biaya tambahan (Rp, opsional)']
];
const TYPE = ['cost', 'benefit', 'cost', 'cost']; // C1 cost, C2 benefit, C3 cost, C4 cost (Modul IV)

/* ---------- Pembacaan file ---------- */
const key = s => String(s).toLowerCase().replace(/[^a-z0-9]/g, '');
const pick = (r, al) => { const k = Object.keys(r).find(k => al.some(a => key(k).includes(a))); return k === undefined ? undefined : r[k]; };
function num(v) {
  if (typeof v === 'number') return v;
  let s = String(v ?? '').replace(/rp|\s|%/gi, ''); if (!s) return NaN;
  const c = s.includes(','), d = s.includes('.');
  if (c && d) s = s.lastIndexOf(',') > s.lastIndexOf('.') ? s.replace(/\./g, '').replace(',', '.') : s.replace(/,/g, '');
  else if (c) s = s.replace(',', '.');
  else if (d && /^-?\d{1,3}(\.\d{3})+$/.test(s)) s = s.replace(/\./g, '');
  return parseFloat(s);
}
function parseCSV(t) {
  const first = t.split(/\r?\n/)[0], cnt = d => first.split(d).length;
  const d = [';', '\t', ','].sort((a, b) => cnt(b) - cnt(a))[0];
  const a = t.trim().split(/\r?\n/).map(l => l.split(d).map(c => c.trim().replace(/^"|"$/g, '')));
  return a.slice(1).filter(r => r.some(c => c !== '')).map(r => Object.fromEntries(a[0].map((k, i) => [k, r[i] ?? ''])));
}
async function readFile(f) {
  if (/\.(csv|txt)$/i.test(f.name)) return parseCSV(await f.text());
  const wb = XLSX.read(await f.arrayBuffer());
  return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '' });
}

/* ---------- Analisis 1: Produksi & mutu (Modul I-II) ---------- */
function runProd(rows) {
  const cm = .75 * (P.std - P.vStd) + .25 * (P.cus - P.vCus), fx = P.fixed + P.order;
  const bep = Math.ceil(fx / cm), bepA = Math.ceil((fx + P.angsuran) / cm);
  const d = rows.map((r, i) => {
    const s = num(pick(r, ['standard', 'std'])) || 0, c = num(pick(r, ['custom', 'cus'])) || 0, x = num(pick(r, ['cacat', 'defect', 'reject'])) || 0;
    const t = s + c, g = t ? (t - x) / t : 0, pc = t ? x / t * 100 : 0;
    const laba = s * g * P.std + c * g * P.cus - s * P.vStd - c * P.vCus - fx;
    const sv = t >= P.target ? 2 : t >= bepA ? 1 : 0, sd = pc <= P.warnCacat ? 2 : pc <= P.cacat ? 1 : 0;
    const why = [sv === 1 ? `Di bawah target ${P.target} pcs` : sv === 0 ? `Di bawah BEP angsuran (${bepA} pcs): angsuran tidak tertutup` : '',
      sd === 1 ? 'Cacat mendekati batas' : sd === 0 ? `Cacat melebihi ${P.cacat}%: hentikan dan perbaiki (SOP-08)` : ''].filter(Boolean).join('; ') || 'Sesuai target';
    return { l: pick(r, ['periode', 'bulan', 'minggu', 'label']) || 'P' + (i + 1), t, pc, laba, st: Math.min(sv, sd), why };
  });
  const avg = k => d.reduce((a, b) => a + b[k], 0) / d.length, last = d[d.length - 1], n0 = d.filter(x => x.st === 0).length;
  const line = (label, v, color, dash) => ({ type: 'line', label, data: d.map(() => v), borderColor: color, borderDash: dash || [6, 4], pointRadius: 0, borderWidth: 2 });
  return {
    st: last.st, brief: `${last.l}: ${fmt(last.t)} pcs, cacat ${fmt(last.pc, 1)}%`,
    msg: `Periode terakhir (${last.l}) berstatus ${L[last.st].toLowerCase()}. ${n0} dari ${d.length} periode tidak aman. BEP ${bep} pcs, BEP dengan angsuran ${bepA} pcs (sudah termasuk biaya pemesanan Modul III).`,
    kpis: [['Rata-rata produksi', avg('t'), 0, '', ' pcs'], ['Rata-rata cacat', avg('pc'), 1, '', '%'], ['Rata-rata laba', avg('laba'), 0, 'Rp', ''], ['Periode aman', d.filter(x => x.st === 2).length, 0, '', ' / ' + d.length]],
    charts: [
      { t: 'Produksi terhadap target dan titik impas', cfg: { type: 'bar', data: { labels: d.map(x => x.l), datasets: [
        { label: 'Produksi (pcs)', data: d.map(x => x.t), backgroundColor: d.map(x => COL[x.st]), borderRadius: 6 },
        line(`Target ${P.target}`, P.target, '#0e7c86'), line(`BEP angsuran ${bepA}`, bepA, '#c98f00'), line(`BEP ${bep}`, bep, '#d64545', [2, 3])] } } },
      { t: 'Persentase cacat terhadap batas', cfg: { type: 'line', data: { labels: d.map(x => x.l), datasets: [
        { label: 'Cacat (%)', data: d.map(x => +x.pc.toFixed(2)), borderColor: '#10252c', backgroundColor: '#10252c', pointBackgroundColor: d.map(x => COL[x.pc <= P.warnCacat ? 2 : x.pc <= P.cacat ? 1 : 0]), pointRadius: 7, tension: .3 },
        line(`Batas ${P.cacat}%`, P.cacat, '#d64545')] } } },
      donut(d)
    ],
    head: ['Periode', 'Total (pcs)', 'Cacat (%)', 'Laba (Rp)', 'Status', 'Alasan'],
    rows: d.map(x => [x.l, fmt(x.t), fmt(x.pc, 1), fmt(x.laba), badge(x.st), x.why])
  };
}

/* ---------- Analisis 2: Persediaan (Modul III) ---------- */
function runStok(rows) {
  const g = {};
  rows.forEach(r => {
    const b = pick(r, ['bahan', 'material']) || 'Paket bahan', a = g[b] = g[b] || [];
    a.push({ l: pick(r, ['minggu', 'periode', 'bulan', 'label']) || 'M' + (a.length + 1), v: num(pick(r, ['stok', 'stock', 'persediaan'])) || 0 });
  });
  const stat = v => v > P.rop ? 2 : v >= P.ss ? 1 : 0, names = Object.keys(g);
  const act = ['Darurat: beli lokal jumlah kecil (SOP-06B)', 'Segera pesan ulang (stok ≤ ROP)', 'Stok cukup'];
  const all = names.flatMap(n => g[n].map(x => ({ ...x, n, st: stat(x.v) })));
  const lasts = names.map(n => g[n][g[n].length - 1]), worst = Math.min(...lasts.map(x => stat(x.v)));
  const pal = ['#0e7c86', '#7a4fd6', '#e0701f', '#2b6cb0', '#a0366b'];
  const ref = (label, v, color) => ({ label, data: g[names[0]].map(() => v), borderColor: color, borderDash: [6, 4], pointRadius: 0, borderWidth: 2 });
  return {
    st: worst, brief: lasts.map((x, i) => `${names[i]}: ${fmt(x.v)} pin`).join(', ').slice(0, 60),
    msg: `Stok terakhir ${worst === 2 ? 'di atas' : worst === 1 ? 'sudah menyentuh' : 'berada di bawah'} batas (ROP ${P.rop} pin, safety stock ${P.ss} pin). Pada MRP Modul III, safety stock baru terbentuk di pesanan kedua, jadi stok rendah di awal bisa wajar.`,
    kpis: [['Stok terakhir', Math.min(...lasts.map(x => x.v)), 0, '', ' pin'], ['Reorder point', P.rop, 0, '', ' pin'], ['Safety stock', P.ss, 0, '', ' pin'], ['Periode darurat', all.filter(x => x.st === 0).length, 0, '', ' / ' + all.length]],
    charts: [
      { t: 'Stok terhadap ROP dan safety stock', cfg: { type: 'line', data: { labels: g[names[0]].map(x => x.l), datasets: [
        ...names.map((n, i) => ({ label: n, data: g[n].map(x => x.v), borderColor: pal[i % 5], backgroundColor: pal[i % 5], pointBackgroundColor: g[n].map(x => COL[stat(x.v)]), pointRadius: 7, tension: .25 })),
        ref(`ROP ${P.rop}`, P.rop, '#c98f00'), ref(`Safety stock ${P.ss}`, P.ss, '#d64545')] } } },
      donut(all)
    ],
    head: ['Periode', 'Bahan', 'Stok (pin)', 'Status', 'Tindakan'],
    rows: all.map(x => [x.l, x.n, fmt(x.v), badge(x.st), act[x.st]])
  };
}

/* ---------- Analisis 3: SAW (Modul IV) ---------- */
function runSaw(rows) {
  const A = rows.map((r, i) => ({ n: pick(r, ['alternatif', 'alt', 'nama']) || 'A' + (i + 1),
    x: [['penyelesaian', 'c1'], ['throughput', 'c2'], ['tunggu', 'c3'], ['biaya', 'c4']].map(a => num(pick(r, a))) }));
  if (A.some(a => a.x.some(isNaN))) throw new Error('kolom');
  const ws = P.w.reduce((a, b) => a + b, 0) || 1;
  A.forEach(a => {
    a.r = a.x.map((v, j) => { const col = A.map(b => b.x[j]); return TYPE[j] === 'benefit' ? v / Math.max(...col) : Math.min(...col) / v; });
    a.V = a.r.reduce((s, r, j) => s + P.w[j] / ws * r, 0);
    a.st = a.V >= P.aman ? 2 : a.V >= P.sedang ? 1 : 0;
    a.over = P.budget !== '' && a.x[3] > num(P.budget);
    if (a.over) a.st = Math.max(0, a.st - 1);
  });
  const S = [...A].sort((a, b) => b.V - a.V), best = S[0], base = Math.min(...A.map(a => a.x[3]));
  const cm = .75 * (P.std - P.vStd) + .25 * (P.cus - P.vCus), labaSetelah = P.target * cm - P.fixed - P.order - P.angsuran;
  const extra = best.x[3] - base;
  const fin = extra > labaSetelah ? ` Perhatian: tambahan biaya ${best.n} (Rp${fmt(extra)}/bulan) lebih besar daripada laba bulanan setelah angsuran (±Rp${fmt(labaSetelah)}). Modul menyarankan meninjau kemampuan dana dan mempertimbangkan alternatif bertahap.` : '';
  const pal = ['#0e7c86', '#e0a100', '#7a4fd6', '#d64545', '#2b6cb0', '#a0366b'];
  return {
    st: best.st, brief: `${best.n} peringkat 1 (V = ${fmt(best.V, 4)})`,
    msg: `Peringkat 1: ${best.n} dengan nilai preferensi V = ${fmt(best.V, 4)}.${best.over ? ' Melebihi batas biaya tambahan sehingga status diturunkan.' : ''}${fin}`,
    kpis: [['Nilai V tertinggi', best.V, 4, '', ''], ['Alternatif dinilai', A.length, 0, '', ''], ['Total bobot', ws, 0, '', '%'], ['Alternatif aman', A.filter(a => a.st === 2).length, 0, '', ' / ' + A.length]],
    charts: [
      { t: 'Nilai preferensi (V) per alternatif', cfg: { type: 'bar', data: { labels: A.map(a => a.n), datasets: [{ label: 'Nilai V', data: A.map(a => +a.V.toFixed(4)), backgroundColor: A.map(a => COL[a.st]), borderRadius: 6 }] }, options: { indexAxis: 'y', scales: { x: { min: 0, max: 1 } } } } },
      { t: 'Nilai ternormalisasi per kriteria', cfg: { type: 'radar', data: { labels: ['C1 Waktu', 'C2 Throughput', 'C3 Tunggu', 'C4 Biaya'], datasets: A.map((a, i) => ({ label: a.n, data: a.r.map(v => +v.toFixed(4)), borderColor: pal[i % 6], backgroundColor: pal[i % 6] + '30' })) }, options: { scales: { r: { min: 0, max: 1 } } } } }
    ],
    head: ['Peringkat', 'Alternatif', 'C1 Waktu (mnt)', 'C2 Throughput', 'C3 Tunggu (mnt)', 'C4 Biaya (Rp)', 'r1', 'r2', 'r3', 'r4', 'V', 'Status'],
    rows: S.map((a, i) => [i + 1, a.n, fmt(a.x[0], 2), fmt(a.x[1], 2), fmt(a.x[2], 2), fmt(a.x[3]), ...a.r.map(v => fmt(v, 4)), `<b>${fmt(a.V, 4)}</b>`, badge(a.st) + (a.over ? ' melebihi anggaran' : '')])
  };
}

/* ---------- Komponen tampilan ---------- */
const badge = s => `<span class="badge ${K[s]}">${L[s]}</span>`;
const donut = list => ({ t: 'Sebaran status', cfg: { type: 'doughnut', data: { labels: L, datasets: [{ data: [0, 1, 2].map(s => list.filter(x => x.st === s).length), backgroundColor: COL }] } } });
const SAMPLE = {
  prod: 'periode,standard,custom,cacat\nBulan 1,110,30,8\nBulan 2,135,45,6\nBulan 3,150,50,7\nBulan 4,155,55,5\nBulan 5,150,50,9\nBulan 6,128,40,4',
  stok: 'minggu,bahan,stok\n' + [125, 50, 25, 0, 176, 101, 76, 51].map((v, i) => `Minggu ${i + 1},Paket bahan,${v}`).join('\n'),
  saw: 'alternatif,waktu_penyelesaian,throughput,waktu_tunggu,biaya_tambahan\nA1 (Kondisi Awal),608.5,19.72,298.5,100000\nA2 (+1 Mesin),409.5,29.30,199,725000\nA3 (+1 Mesin & +1 Pekerja),310,38.71,149.25,2225000'
};
const MODS = {
  prod: { t: 'Produksi & Mutu', s: 'Modul I–II', hint: 'Kolom: periode, standard, custom, cacat (pcs)', run: runProd },
  stok: { t: 'Persediaan', s: 'Modul III', hint: 'Kolom: minggu, bahan, stok (satuan pin). Data contoh diambil dari tabel MRP Modul III.', run: runStok },
  saw: { t: 'Pilih Alternatif (SAW)', s: 'Modul IV', hint: 'Kolom: alternatif, waktu_penyelesaian, throughput, waktu_tunggu, biaya_tambahan. Data contoh = Tabel 2.1 Modul IV.', run: runSaw }
};
const data = {}, res = {}, charts = {};

Chart.defaults.font.family = 'Figtree, system-ui, sans-serif';
Chart.defaults.animation = { duration: reduce ? 0 : 1300, easing: 'easeOutQuart' };
Chart.defaults.maintainAspectRatio = true;
Chart.defaults.plugins.legend.position = 'bottom';

function build() {
  const tabs = $('#tabs'), panels = $('#panels');
  Object.entries(MODS).forEach(([id, m], i) => {
    tabs.insertAdjacentHTML('beforeend', `<button role="tab" data-id="${id}" aria-selected="${i === 0}">${m.t} <small>(${m.s})</small></button>`);
    panels.insertAdjacentHTML('beforeend', `<section class="panel ${i === 0 ? 'on' : ''}" id="p-${id}" role="tabpanel">
      <div class="drop" data-id="${id}"><div><p><b>Seret file ke sini</b> atau pilih dari komputer (.csv, .xlsx)</p><small>${m.hint}</small></div>
      <div class="btns"><label class="btn">Pilih file<input type="file" accept=".csv,.xlsx,.xls,.txt" hidden></label>
      <button class="btn alt" data-act="sample">Pakai data contoh</button><button class="btn alt" data-act="tpl">Unduh template</button></div></div>
      <div id="out-${id}"></div></section>`);
  });
  tabs.onclick = e => { const b = e.target.closest('button'); if (!b) return;
    tabs.querySelectorAll('button').forEach(x => x.setAttribute('aria-selected', x === b));
    document.querySelectorAll('.panel').forEach(p => p.classList.toggle('on', p.id === 'p-' + b.dataset.id)); };
  document.querySelectorAll('.drop').forEach(z => {
    const id = z.dataset.id;
    z.querySelector('input').onchange = e => e.target.files[0] && fromFile(id, e.target.files[0]);
    z.onclick = e => { const a = e.target.dataset.act; if (a === 'sample') load(id, parseCSV(SAMPLE[id])); if (a === 'tpl') download(id); };
    ['dragover', 'dragleave', 'drop'].forEach(ev => z.addEventListener(ev, e => { e.preventDefault(); z.classList.toggle('over', ev === 'dragover'); if (ev === 'drop' && e.dataTransfer.files[0]) fromFile(id, e.dataTransfer.files[0]); }));
  });
  $('#pgrid').innerHTML = PARAMS.map(([k, l]) => `<label>${l}<input data-k="${k}" type="number" step="any" value="${k[0] === 'w' && k.length === 2 ? P.w[+k[1]] : P[k]}"></label>`).join('');
  $('#pgrid').oninput = e => { const k = e.target.dataset.k, v = e.target.value;
    if (k[0] === 'w' && k.length === 2) P.w[+k[1]] = +v || 0; else P[k] = k === 'budget' ? v : +v;
    Object.keys(data).forEach(id => load(id, data[id], true)); };
  summary();
}
async function fromFile(id, f) {
  try { load(id, await readFile(f)); } catch (e) { fail(id); }
}
function download(id) {
  const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([SAMPLE[id]], { type: 'text/csv' }));
  a.download = 'template_' + id + '.csv'; a.click(); URL.revokeObjectURL(a.href);
}
const fail = id => { $('#out-' + id).innerHTML = '<div class="err">Data belum bisa dibaca. Pastikan baris pertama berisi nama kolom seperti pada template dan nilainya berupa angka. Klik “Unduh template” untuk contoh formatnya.</div>'; };
function load(id, rows, quiet) {
  if (!rows.length) return fail(id);
  try { data[id] = rows; res[id] = MODS[id].run(rows); } catch (e) { delete res[id]; summary(); return fail(id); }
  const r = res[id], out = $('#out-' + id);
  (charts[id] || []).forEach(c => c.destroy());
  out.innerHTML = `<div class="verdict ${K[r.st]}"><span class="dot ${quiet ? '' : 'pulse'}"></span><div><b>${L[r.st]}</b><p>${r.msg}</p></div></div>
    <div class="kpis">${r.kpis.map(k => `<div class="kpi"><span class="num" data-v="${k[1]}" data-d="${k[2]}" data-p="${k[3]}">0</span><small>${k[4]}</small><em>${k[0]}</em></div>`).join('')}</div>
    <div class="charts">${r.charts.map((c, i) => `<figure><figcaption>${c.t}</figcaption><canvas id="c-${id}-${i}"></canvas></figure>`).join('')}</div>
    <div class="tw"><table><thead><tr>${r.head.map(h => `<th>${h}</th>`).join('')}</tr></thead><tbody>${r.rows.map(row => `<tr>${row.map(c => `<td>${c}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`;
  charts[id] = r.charts.map((c, i) => new Chart($(`#c-${id}-${i}`), c.cfg));
  out.querySelectorAll('.num').forEach(countUp);
  summary();
}
function countUp(el) {
  const v = +el.dataset.v, d = +el.dataset.d, p = el.dataset.p, t0 = performance.now();
  const set = x => el.textContent = (p || '') + fmt(x, d);
  if (reduce) return set(v);
  (function f(t) { const k = Math.min(1, (t - t0) / 1000); set(v * (1 - Math.pow(1 - k, 3))); if (k < 1) requestAnimationFrame(f); })(t0);
}
function summary() {
  Object.keys(MODS).forEach(id => {
    const r = res[id], c = $('#sum-' + id);
    c.className = 'sc ' + (r ? K[r.st] : '');
    c.querySelector('b').textContent = r ? L[r.st] : 'Belum ada data';
    c.querySelector('p').textContent = r ? r.brief : 'Unggah file atau pakai data contoh';
  });
  const all = Object.values(res), o = $('#overall');
  if (!all.length) return;
  const m = Math.min(...all.map(r => r.st));
  o.className = 'overall ' + K[m];
  o.querySelector('b').textContent = 'Keputusan: ' + L[m];
  o.querySelector('p').textContent = all.length < 3 ? `Berdasarkan ${all.length} dari 3 modul yang sudah diunggah` : 'Status terendah dari ketiga modul menentukan keputusan';
}
build();
