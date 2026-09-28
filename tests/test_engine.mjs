// Tes engine Fuzzy DEMATEL (CLAUDE.md §8).
// Jalankan: npm install && node tests/test_engine.mjs
//
// Mengambil isi <script id="engine"> dari fuzzy_dematel.html, menjalankannya di Node,
// lalu membandingkan hasilnya dengan nilai acuan dari implementasi Python (NumPy) independen.

import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';
import XLSX from 'xlsx';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const HTML_PATH = path.join(ROOT, 'fuzzy_dematel.html');
const DATA_PATH = path.join(ROOT, 'Responden_final_thesis.xlsx');

// ---------- muat engine dari HTML ----------
const html = fs.readFileSync(HTML_PATH, 'utf8');
const match = html.match(/<script id="engine">([\s\S]*?)<\/script>/);
if (!match) throw new Error('<script id="engine"> tidak ditemukan di fuzzy_dematel.html');
const context = {};
vm.createContext(context);
vm.runInContext(match[1], context, { filename: 'engine.js' });
const E = context.FuzzyDematelEngine;

// ---------- harness kecil ----------
let passed = 0, failed = 0;
const failures = [];
function test(name, fn) {
  try { fn(); passed++; console.log('  ✓ ' + name); }
  catch (err) { failed++; failures.push(name); console.log('  ✗ ' + name + '\n      ' + err.message.replace(/\n/g, '\n      ')); }
}
function section(title) { console.log('\n' + title); }
function assert(cond, msg) { if (!cond) throw new Error(msg); }
function close(actual, expected, tol, label) {
  if (!(Math.abs(actual - expected) <= tol)) {
    throw new Error(`${label}: dapat ${actual}, acuan ${expected} (selisih ${Math.abs(actual - expected).toExponential(2)} > ${tol})`);
  }
}
function includesSome(list, parts, label) {
  const hit = list.find((m) => parts.every((p) => m.includes(p)));
  assert(hit, `${label}: tidak ada pesan yang memuat ${JSON.stringify(parts)}.\n  Pesan: ${JSON.stringify(list.slice(0, 10))}`);
  return hit;
}

// Tulis workbook ke bytes lalu baca ulang, meniru file .xlsx yang diunggah pengguna.
function roundTrip(wb) {
  return XLSX.read(XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }), { type: 'buffer' });
}
function loadOriginal() { return XLSX.readFile(DATA_PATH); }
function setCell(wb, sheet, addr, value) {
  const ws = wb.Sheets[sheet];
  if (value === undefined) { delete ws[addr]; return; }
  ws[addr] = typeof value === 'number' ? { t: 'n', v: value } : { t: 's', v: value };
}

// ---------- nilai acuan (CLAUDE.md §8) ----------
const REFERENCE = {
  row_influences_column: {
    scale_factor_s: 15.985714, cfcs_min_l: 0.015107, cfcs_max_u: 0.719461,
    threshold_mean_offdiag: 0.209336, threshold_mean_all: 0.20756, n_significant_with_offdiag_threshold: 196,
    Z_agg_R1_to_R2: [0.557143, 0.757143, 0.914286], T_fuzzy_R1_to_R2: [0.060836, 0.145453, 0.699541], T_crisp_R1_to_R2: 0.23669,
    factors: {
      R1: [4.5049, 4.239, 8.7439, 0.266], R2: [4.2778, 4.258, 8.5359, 0.0198], R3: [4.1835, 4.242, 8.4255, -0.0585],
      R4: [4.2736, 4.0479, 8.3215, 0.2257], R5: [4.1509, 4.175, 8.3259, -0.0241], R6: [4.0553, 4.1026, 8.1579, -0.0473],
      R7: [4.2301, 4.3804, 8.6106, -0.1503], R8: [3.9446, 3.8443, 7.7889, 0.1003], R9: [4.1143, 4.1222, 8.2364, -0.0079],
      R10: [4.2428, 4.1732, 8.4161, 0.0696], R11: [4.1248, 4.2331, 8.3579, -0.1083], R12: [4.2381, 4.1808, 8.4188, 0.0573],
      R13: [4.1102, 4.1466, 8.2568, -0.0364], R14: [3.7386, 3.7729, 7.5115, -0.0342], R15: [4.049, 4.0521, 8.1011, -0.0031],
      R16: [4.2423, 4.2891, 8.5313, -0.0468], R17: [3.9781, 4.0752, 8.0533, -0.097], R18: [4.1478, 4.1644, 8.3122, -0.0166],
      R19: [4.2332, 4.2663, 8.4995, -0.0331], R20: [4.1842, 4.2591, 8.4432, -0.0749]
    }
  },
  column_influences_row: {
    scale_factor_s: 15.628571, cfcs_min_l: 0.016133, cfcs_max_u: 1.033155,
    threshold_mean_offdiag: 0.266823, threshold_mean_all: 0.264972, n_significant_with_offdiag_threshold: 199,
    Z_agg_R1_to_R2: [0.557143, 0.757143, 0.9], T_fuzzy_R1_to_R2: [0.061048, 0.150155, 0.948403], T_crisp_R1_to_R2: 0.286556,
    factors: {
      R1: [5.4041, 5.7405, 11.1446, -0.3364], R2: [5.4316, 5.4557, 10.8874, -0.0241], R3: [5.4112, 5.3372, 10.7483, 0.074],
      R4: [5.1709, 5.4516, 10.6225, -0.2807], R5: [5.3285, 5.3009, 10.6294, 0.0276], R6: [5.2381, 5.1796, 10.4177, 0.0585],
      R7: [5.5872, 5.3962, 10.9834, 0.191], R8: [4.9155, 5.041, 9.9564, -0.1255], R9: [5.2608, 5.2508, 10.5117, 0.01],
      R10: [5.3271, 5.4127, 10.7399, -0.0856], R11: [5.4015, 5.2646, 10.6661, 0.1368], R12: [5.3394, 5.4069, 10.7463, -0.0675],
      R13: [5.2931, 5.2472, 10.5403, 0.0458], R14: [4.8287, 4.7842, 9.6129, 0.0445], R15: [5.1777, 5.1751, 10.3528, 0.0026],
      R16: [5.47, 5.4162, 10.8863, 0.0538], R17: [5.2067, 5.0835, 10.2903, 0.1232], R18: [5.3192, 5.2969, 10.616, 0.0223],
      R19: [5.4422, 5.4045, 10.8467, 0.0377], R20: [5.4353, 5.3433, 10.7786, 0.092]
    }
  }
};
const TOL = 1e-4, TOL_RC = 5e-4;

// =====================================================================
section('1. Data asli Responden_final_thesis.xlsx');
const parsedOriginal = E.parseWorkbook(loadOriginal());

test('parse: 7 responden, 20 faktor, tanpa error', () => {
  assert(parsedOriginal.errors.length === 0, 'error: ' + JSON.stringify(parsedOriginal.errors));
  assert(parsedOriginal.experts.length === 7, 'jumlah responden ' + parsedOriginal.experts.length);
  assert(parsedOriginal.codes.length === 20, 'jumlah faktor ' + parsedOriginal.codes.length);
  assert(parsedOriginal.names[1] === 'Trade Route Disruption', 'nama R2 harus di-trim: ' + JSON.stringify(parsedOriginal.names[1]));
});

test('info: peringatan diagonal pada file asli', () => {
  // CLAUDE.md §3 menyebut responden 7 sel R11→R11 berisi 4. Pada file yang diterima, sel itu (M13) bernilai 0,
  // sehingga file asli tidak menghasilkan peringatan. Deteksi peringatannya diuji di bagian 3 dengan file yang dimodifikasi.
  const ws = loadOriginal().Sheets['responden 7'];
  console.log(`      (responden 7!M13 = ${ws.M13 && ws.M13.v}; peringatan pada file asli: ${parsedOriginal.warnings.length})`);
});

for (const orientation of Object.keys(REFERENCE)) {
  const ref = REFERENCE[orientation];
  const r = E.runFuzzyDematel(parsedOriginal.experts, { orientation, thresholdMethod: 'mean_offdiag' });
  test(`[${orientation}] s, min_l, max_u, threshold, jumlah signifikan`, () => {
    close(r.s, ref.scale_factor_s, TOL, 's');
    close(r.minL, ref.cfcs_min_l, TOL, 'min_l');
    close(r.maxU, ref.cfcs_max_u, TOL, 'max_u');
    close(r.threshold.alpha, ref.threshold_mean_offdiag, TOL, 'α mean_offdiag');
    close(r.threshold.meanAll, ref.threshold_mean_all, TOL, 'α mean_all');
    assert(r.significant.length === ref.n_significant_with_offdiag_threshold,
      `jumlah signifikan ${r.significant.length}, acuan ${ref.n_significant_with_offdiag_threshold}`);
  });
  test(`[${orientation}] Z, T fuzzy, T crisp untuk R1 → R2`, () => {
    ['l', 'm', 'u'].forEach((c, k) => {
      close(r.Z[c][0][1], ref.Z_agg_R1_to_R2[k], TOL, 'Z_' + c);
      close(r.T[c][0][1], ref.T_fuzzy_R1_to_R2[k], TOL, 'T_' + c);
    });
    close(r.crisp[0][1], ref.T_crisp_R1_to_R2, TOL, 'T_crisp');
  });
  test(`[${orientation}] R, C, R+C, R−C untuk 20 faktor`, () => {
    parsedOriginal.codes.forEach((code, i) => {
      const [R, C, P, D] = ref.factors[code];
      close(r.rc.R[i], R, TOL_RC, code + ' R');
      close(r.rc.C[i], C, TOL_RC, code + ' C');
      close(r.rc.prominence[i], P, TOL_RC, code + ' R+C');
      close(r.rc.relation[i], D, TOL_RC, code + ' R−C');
      const expected = D > 0 ? 'Cause' : 'Effect';
      assert(r.rc.classification[i] === expected, `${code} klasifikasi ${r.rc.classification[i]}, seharusnya ${expected}`);
    });
  });
  test(`[${orientation}] ranking & hubungan signifikan konsisten`, () => {
    const byP = [...r.rc.prominence.keys()].sort((a, b) => r.rc.prominence[b] - r.rc.prominence[a]);
    byP.forEach((i, pos) => assert(r.rc.rankProminence[i] === pos + 1, 'rank R+C salah untuk ' + parsedOriginal.codes[i]));
    for (let k = 1; k < r.significant.length; k++) assert(r.significant[k - 1].value >= r.significant[k].value, 'daftar tidak menurun');
    assert(r.significant.every((s) => s.from !== s.to && s.value > r.threshold.alpha), 'ada pasangan diagonal atau ≤ α');
  });
}

test('orientasi column_influences_row = transpose tiap responden', () => {
  const a = E.runFuzzyDematel(parsedOriginal.experts, { orientation: 'column_influences_row' });
  const transposed = parsedOriginal.experts.map((e) => ({ sheet: e.sheet, scores: e.scores[0].map((_, j) => e.scores.map((row) => row[j])) }));
  const b = E.runFuzzyDematel(transposed, { orientation: 'row_influences_column' });
  for (let i = 0; i < 20; i++) for (let j = 0; j < 20; j++) close(a.crisp[i][j], b.crisp[i][j], 1e-12, `crisp[${i}][${j}]`);
});

test('threshold mean_all dan manual', () => {
  const all = E.runFuzzyDematel(parsedOriginal.experts, { thresholdMethod: 'mean_all' });
  close(all.threshold.alpha, REFERENCE.row_influences_column.threshold_mean_all, TOL, 'α mean_all');
  const man = E.runFuzzyDematel(parsedOriginal.experts, { thresholdMethod: 'manual', manualThreshold: 0.25 });
  assert(man.threshold.alpha === 0.25, 'α manual');
  assert(man.significant.every((s) => s.value > 0.25), 'signifikan manual');
  let threw = false;
  try { E.runFuzzyDematel(parsedOriginal.experts, { thresholdMethod: 'manual', manualThreshold: NaN }); } catch (e) { threw = true; }
  assert(threw, 'α manual kosong harus error');
});

// =====================================================================
section('2. Validasi (error harus menyebut sheet + sel)');

test('sel kosong, nilai 5, nilai 2.5, dan teks → error dengan sheet & sel', () => {
  const wb = loadOriginal();
  setCell(wb, 'responden 3', 'F12', undefined); // kosong
  setCell(wb, 'responden 2', 'G5', 5);          // di luar skala
  setCell(wb, 'responden 1', 'D9', 2.5);        // bukan bilangan bulat
  setCell(wb, 'responden 4', 'H9', 'abc');      // teks
  const p = E.parseWorkbook(roundTrip(wb));
  includesSome(p.errors, ['responden 3, sel F12', 'kosong'], 'sel kosong');
  includesSome(p.errors, ['responden 2, sel G5', '5', 'di luar'], 'nilai 5');
  includesSome(p.errors, ['responden 1, sel D9', '2.5'], 'nilai 2.5');
  includesSome(p.errors, ['responden 4, sel H9', 'bukan angka'], 'teks');
  assert(p.errors.length === 4, 'jumlah error ' + p.errors.length + ': ' + JSON.stringify(p.errors));
  assert(p.experts.length === 3, 'sheet yang error tidak boleh ikut dihitung');
});

test('kode kolom tidak berurutan → error', () => {
  const wb = loadOriginal();
  setCell(wb, 'responden 5', 'E1', 'R4'); // seharusnya R3
  const p = E.parseWorkbook(roundTrip(wb));
  includesSome(p.errors, ['responden 5, sel E1', 'R3'], 'kode kolom');
});

test('kode baris tidak berurutan → error', () => {
  const wb = loadOriginal();
  setCell(wb, 'responden 6', 'A10', 'R9'); // seharusnya R8
  const p = E.parseWorkbook(roundTrip(wb));
  includesSome(p.errors, ['responden 6, sel A10', 'R8'], 'kode baris');
});

test('tidak ada sheet responden → error', () => {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['Instruction'], ['rate the degree...']]), 'Instruction');
  const p = E.parseWorkbook(roundTrip(wb));
  includesSome(p.errors, ['Tidak ada sheet responden'], 'tanpa responden');
});

test('sheet non-responden (A3 ≠ R1) diabaikan', () => {
  const wb = loadOriginal();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['Ringkasan'], ['x'], ['total', 1, 2]]), 'Ringkasan');
  const p = E.parseWorkbook(roundTrip(wb));
  assert(p.errors.length === 0 && p.experts.length === 7, 'harus tetap 7 responden');
  assert(p.skippedSheets.includes('Ringkasan'), 'Ringkasan harus tercatat diabaikan');
});

// =====================================================================
section('3. Peringatan diagonal');

test('responden 7 sel R11→R11 (M13) = 4 → peringatan kuning, diagonal dipaksa 0, hasil tidak berubah', () => {
  const wb = loadOriginal();
  setCell(wb, 'responden 7', 'M13', 4);
  const p = E.parseWorkbook(roundTrip(wb));
  assert(p.errors.length === 0, 'diagonal tidak boleh jadi error: ' + JSON.stringify(p.errors));
  includesSome(p.warnings, ['responden 7, sel M13', 'R11', 'diagonal'], 'peringatan diagonal');
  assert(p.experts[6].scores[10][10] === 0, 'diagonal harus 0');
  const a = E.runFuzzyDematel(p.experts, {}), b = E.runFuzzyDematel(parsedOriginal.experts, {});
  for (let i = 0; i < 20; i++) for (let j = 0; j < 20; j++) close(a.crisp[i][j], b.crisp[i][j], 1e-15, 'crisp');
});

// =====================================================================
section('4. Tes kecil buatan tangan');

test('invert 2×2 [[4,7],[2,6]] = [[0.6,−0.7],[−0.2,0.4]]', () => {
  const inv = E.invert([[4, 7], [2, 6]]);
  [[0.6, -0.7], [-0.2, 0.4]].forEach((row, i) => row.forEach((v, j) => close(inv[i][j], v, 1e-12, `inv[${i}][${j}]`)));
});

test('invert butuh pivoting: [[0,1],[1,0]] dan matriks singular → error', () => {
  const inv = E.invert([[0, 1], [1, 0]]);
  close(inv[0][1], 1, 1e-15, 'inv[0][1]');
  let msg = '';
  try { E.invert([[1, 2], [2, 4]]); } catch (e) { msg = e.message; }
  assert(/singular/i.test(msg), 'harus error "matriks singular", dapat: ' + msg);
});

test('3×3, 1 pakar: skor [[0,4,2],[1,0,3],[0,0,0]]', () => {
  // Hitung tangan:
  //   TFN baris R1: (0,0,0), (0.7,0.9,1.0), (0.3,0.5,0.7) → jumlah u = 1.7
  //   TFN baris R2: (0.1,0.3,0.5), (0,0,0), (0.5,0.7,0.9)  → jumlah u = 1.4
  //   TFN baris R3: (0,0.1,0.3), (0,0.1,0.3), (0,0,0)      → jumlah u = 0.6
  //   s = 1.7. Baris R3 komponen l = 0 → baris R3 pada T_l = 0, sehingga min_l = 0.
  // Nilai T crisp & R/C dibandingkan dengan NumPy.
  const r = E.runFuzzyDematel([{ sheet: 'tangan', scores: [[0, 4, 2], [1, 0, 3], [0, 0, 0]] }], {});
  close(r.s, 1.7, 1e-12, 's');
  close(r.Z.u[0][1], 1.0, 1e-12, 'Z_u R1→R2');
  close(r.X.m[1][2], 0.7 / 1.7, 1e-12, 'X_m R2→R3');
  [0, 1, 2].forEach((j) => close(r.T.l[2][j], 0, 1e-15, 'T_l baris R3'));
  close(r.minL, 0, 1e-15, 'min_l');
  close(r.maxU, 1.237116991643454, 1e-12, 'max_u');
  close(r.T.l[0][1], 0.4219858156028369, 1e-12, 'T_l R1→R2');
  const crisp = [[0.2122942529834524, 0.6892450568362042, 0.6590043009837566],
    [0.299475216194075, 0.225470570872751, 0.6097159142304132],
    [0.1265492086510393, 0.16369911906823192, 0.1202851572676371]];
  crisp.forEach((row, i) => row.forEach((v, j) => close(r.crisp[i][j], v, 1e-12, `crisp[${i}][${j}]`)));
  [1.5605436108034132, 1.1346617012972393, 0.41053348498690834].forEach((v, i) => close(r.rc.R[i], v, 1e-12, 'R' + (i + 1)));
  [0.6383186778285668, 1.0784147467771872, 1.3890053724818068].forEach((v, i) => close(r.rc.C[i], v, 1e-12, 'C' + (i + 1)));
  assert(r.rc.classification.join(',') === 'Cause,Cause,Effect', 'klasifikasi ' + r.rc.classification.join(','));
  close(r.threshold.alpha, 0.42461480266062007, 1e-12, 'α');
  assert(r.significant.length === 3, 'jumlah signifikan ' + r.significant.length);
});

test('CFCS satu sel: TFN simetris (0.2, 0.5, 0.8) dengan min_l=0, Δ=1 → 0.5', () => {
  close(E.cfcsCell(0.2, 0.5, 0.8, 0, 1).crisp, 0.5, 1e-12, 'crisp');
});

// =====================================================================
console.log(`\n${passed} lulus, ${failed} gagal`);
if (failed) { console.log('Gagal: ' + failures.join('; ')); process.exit(1); }
