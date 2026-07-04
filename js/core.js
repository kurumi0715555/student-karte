/* 生徒カルテ core — データモデル・Excel入出力・保存機構
   正本は常にExcelファイル。ブラウザ側にデータを残さない。 */
'use strict';

/* ========== 定数 ========== */
const SHEET_MEIBO = '名簿';
const SHEET_RECORDS = '記録';
const SHEET_RELATIONS = '関係';
const SHEET_GRADES = '成績';

const MEIBO_COLS = [
  '生徒ID', '生徒氏名', 'ふりがな', '性別', '生年月日', '出身小学校',
  '住所', '連絡先番号', '連絡先氏名', '兄弟姉妹', '地区・通学', '健康メモ',
  'クラブ', '委員会', '支援計画',
  '1年クラス', '1年番号', '2年クラス', '2年番号', '3年クラス', '3年番号',
  '在籍', '備考'
];
const SENSITIVE_COLS = ['住所', '連絡先番号', '連絡先氏名', '健康メモ', '支援計画'];
const RECORD_COLS = ['日付', '生徒ID', '氏名', '種別', '内容', '記録者'];
const RELATION_COLS = ['生徒ID_A', '氏名A', '生徒ID_B', '氏名B', '種別', '理由・メモ'];
const GRADE_COLS = ['生徒ID', '氏名', '年度', '学期', 'テスト名', '教科', '点数'];

const RECORD_TYPES = ['指導', '支援', '面談', '保護者対応', '健康', 'クラブ', 'その他'];
const RELATION_TYPES = ['同クラス不可', '同クラス推奨', '要注意', 'その他'];
const STATUS_ACTIVE = '在籍';
const FILE_DEFAULT_NAME = '生徒カルテ.xlsx';

/* ========== 状態 ========== */
const state = {
  loaded: false,
  meibo: [],            // 1行1生徒のオブジェクト配列（列名キー・すべて文字列）
  meiboHeaders: [],
  records: [],
  recordsHeaders: [],
  relations: [],
  relationsHeaders: [],
  grades: [],
  gradesHeaders: [],
  extraSheets: [],      // {name, ws} 未知シートはそのまま書き戻す
  fileHandle: null,
  fileName: FILE_DEFAULT_NAME,
  saveMode: null,       // 'fsa' | 'download'
  dirty: 0,
  saveError: false,
  nendo: 0              // 表示中の年度（令和）
};

/* ========== 年度・ID ========== */
function currentNendo() {
  const now = new Date();
  const fy = now.getFullYear() - (now.getMonth() + 1 < 4 ? 1 : 0);
  return fy - 2018; // 2019年度 = 令和1年度
}

function entryYearOf(id) {
  const s = String(id || '').trim();
  if (!/^\d{4,5}$/.test(s)) return NaN;
  return parseInt(s.slice(0, s.length - 3), 10);
}

function gradeOf(student, nendo) {
  const e = entryYearOf(student['生徒ID']);
  if (isNaN(e)) return NaN;
  return (nendo || state.nendo) - e + 1;
}

function classColOf(grade) { return grade + '年クラス'; }
function numColOf(grade) { return grade + '年番号'; }

function currentClass(student, nendo) {
  const g = gradeOf(student, nendo);
  return (g >= 1 && g <= 3) ? String(student[classColOf(g)] || '') : '';
}
function currentNumber(student, nendo) {
  const g = gradeOf(student, nendo);
  return (g >= 1 && g <= 3) ? String(student[numColOf(g)] || '') : '';
}

function nextIdFor(entryYear) {
  let maxSeq = 0;
  for (const s of state.meibo) {
    const id = String(s['生徒ID'] || '');
    if (entryYearOf(id) === entryYear) {
      maxSeq = Math.max(maxSeq, parseInt(id.slice(-3), 10) || 0);
    }
  }
  return String(entryYear) + String(maxSeq + 1).padStart(3, '0');
}

/* ========== 文字処理 ========== */
function normName(s) {
  return String(s || '').replace(/[\s　]/g, '').trim();
}

function parseTSV(text) {
  return String(text || '')
    .split(/\r?\n/)
    .map(line => line.split('\t').map(c => c.trim()))
    .filter(cells => cells.some(c => c !== ''));
}

function todayStr() {
  const d = new Date();
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
}

function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, c =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

/* ========== Excel 読み込み ========== */
function sheetToRows(ws) {
  if (!ws) return { headers: [], rows: [] };
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });
  if (aoa.length === 0) return { headers: [], rows: [] };
  const headers = aoa[0].map(h => String(h).trim()).filter(h => h !== '');
  const rows = [];
  for (let i = 1; i < aoa.length; i++) {
    if (aoa[i].every(c => String(c).trim() === '')) continue;
    const obj = {};
    headers.forEach((h, j) => { obj[h] = String(aoa[i][j] ?? '').trim(); });
    rows.push(obj);
  }
  return { headers, rows };
}

function mergeHeaders(found, required) {
  // 既存の列順を尊重しつつ、必須列で欠けているものを末尾に足す
  const out = found.slice();
  for (const r of required) if (!out.includes(r)) out.push(r);
  return out;
}

function loadWorkbook(arrayBuffer, fileName) {
  const wb = XLSX.read(arrayBuffer, { type: 'array' });
  const problems = { errors: [], warnings: [] };

  if (!wb.SheetNames.includes(SHEET_MEIBO)) {
    problems.errors.push('シート「名簿」が見つかりません。テンプレートから作成したファイルを開いてください。');
    return { problems };
  }

  const meibo = sheetToRows(wb.Sheets[SHEET_MEIBO]);
  if (!meibo.headers.includes('生徒ID') || !meibo.headers.includes('生徒氏名')) {
    problems.errors.push('名簿シートに「生徒ID」「生徒氏名」の列が必要です。');
    return { problems };
  }

  const records = sheetToRows(wb.Sheets[SHEET_RECORDS]);
  const relations = sheetToRows(wb.Sheets[SHEET_RELATIONS]);
  const grades = sheetToRows(wb.Sheets[SHEET_GRADES]);

  // 検証
  const seen = new Map();
  meibo.rows.forEach((r, i) => {
    const id = String(r['生徒ID'] || '').trim();
    if (id === '') {
      problems.warnings.push(`名簿 ${i + 2}行目: 生徒IDが空欄です（${r['生徒氏名'] || '氏名不明'}）`);
    } else {
      if (seen.has(id)) problems.warnings.push(`生徒IDが重複しています: ${id}（${seen.get(id)} / ${r['生徒氏名']}）`);
      seen.set(id, r['生徒氏名']);
      if (isNaN(entryYearOf(id))) problems.warnings.push(`生徒IDの形式が不正です: ${id}（入学年度+連番3桁の数字）`);
    }
  });

  state.meibo = meibo.rows;
  state.meiboHeaders = mergeHeaders(meibo.headers, MEIBO_COLS);
  state.records = records.rows;
  state.recordsHeaders = mergeHeaders(records.headers, RECORD_COLS);
  state.relations = relations.rows;
  state.relationsHeaders = mergeHeaders(relations.headers, RELATION_COLS);
  state.grades = grades.rows;
  state.gradesHeaders = mergeHeaders(grades.headers, GRADE_COLS);

  const known = [SHEET_MEIBO, SHEET_RECORDS, SHEET_RELATIONS, SHEET_GRADES];
  state.extraSheets = wb.SheetNames.filter(n => !known.includes(n)).map(n => ({ name: n, ws: wb.Sheets[n] }));

  state.fileName = fileName || FILE_DEFAULT_NAME;
  state.loaded = true;
  state.dirty = 0;
  state.saveError = false;
  return { problems };
}

/* ========== Excel 書き出し ========== */
function refreshRefNames() {
  const nameOf = new Map(state.meibo.map(s => [String(s['生徒ID']), String(s['生徒氏名'] || '')]));
  for (const r of state.records) if (nameOf.has(String(r['生徒ID']))) r['氏名'] = nameOf.get(String(r['生徒ID']));
  for (const r of state.grades) if (nameOf.has(String(r['生徒ID']))) r['氏名'] = nameOf.get(String(r['生徒ID']));
  for (const r of state.relations) {
    if (nameOf.has(String(r['生徒ID_A']))) r['氏名A'] = nameOf.get(String(r['生徒ID_A']));
    if (nameOf.has(String(r['生徒ID_B']))) r['氏名B'] = nameOf.get(String(r['生徒ID_B']));
  }
}

function rowsToSheet(rows, headers) {
  const ws = XLSX.utils.json_to_sheet(rows, { header: headers, skipHeader: false });
  ws['!cols'] = headers.map(h => ({ wch: Math.max(10, h.length * 2 + 4) }));
  return ws;
}

function buildWorkbook() {
  refreshRefNames();
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, rowsToSheet(state.meibo, state.meiboHeaders), SHEET_MEIBO);
  XLSX.utils.book_append_sheet(wb, rowsToSheet(state.records, state.recordsHeaders), SHEET_RECORDS);
  XLSX.utils.book_append_sheet(wb, rowsToSheet(state.relations, state.relationsHeaders), SHEET_RELATIONS);
  XLSX.utils.book_append_sheet(wb, rowsToSheet(state.grades, state.gradesHeaders), SHEET_GRADES);
  for (const ex of state.extraSheets) XLSX.utils.book_append_sheet(wb, ex.ws, ex.name);
  return wb;
}

function workbookToBlob(wb) {
  const data = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
  return new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

function downloadBlob(blob, filename) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
}

/* ========== 保存 ========== */
const fsaSupported = typeof window.showOpenFilePicker === 'function';

async function saveToHandle() {
  const blob = workbookToBlob(buildWorkbook());
  const writable = await state.fileHandle.createWritable();
  await writable.write(blob);
  await writable.close();
}

async function persist() {
  // 変更確定のたびに呼ぶ。fsaなら即時上書き、downloadなら未保存カウント。
  state.dirty++;
  if (state.saveMode === 'fsa') {
    try {
      await saveToHandle();
      state.dirty = 0;
      state.saveError = false;
    } catch (e) {
      console.error('保存に失敗:', e);
      state.saveError = true;
    }
  }
  if (typeof onSaveStateChanged === 'function') onSaveStateChanged();
}

async function retrySave() {
  if (state.saveMode !== 'fsa') return;
  try {
    await saveToHandle();
    state.dirty = 0;
    state.saveError = false;
  } catch (e) {
    console.error('再保存に失敗:', e);
    state.saveError = true;
  }
  if (typeof onSaveStateChanged === 'function') onSaveStateChanged();
}

function manualDownloadSave() {
  downloadBlob(workbookToBlob(buildWorkbook()), state.fileName);
  state.dirty = 0;
  state.saveError = false;
  if (typeof onSaveStateChanged === 'function') onSaveStateChanged();
}

function downloadBackup() {
  const base = state.fileName.replace(/\.xlsx$/i, '');
  downloadBlob(workbookToBlob(buildWorkbook()), `${base}_backup_${todayStr()}.xlsx`);
}

/* ========== ファイルを開く ========== */
async function openWithPicker() {
  const [handle] = await window.showOpenFilePicker({
    types: [{ description: 'Excelブック', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] } }],
    multiple: false
  });
  const perm = await handle.requestPermission({ mode: 'readwrite' });
  const file = await handle.getFile();
  const buf = await file.arrayBuffer();
  const result = loadWorkbook(buf, file.name);
  if (result.problems.errors.length === 0) {
    if (perm === 'granted') {
      state.fileHandle = handle;
      state.saveMode = 'fsa';
    } else {
      state.fileHandle = null;
      state.saveMode = 'download';
      result.problems.warnings.push('書き込み許可が得られなかったため、保存はダウンロード方式になります。');
    }
  }
  return result;
}

function openWithInput() {
  return new Promise((resolve, reject) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.xlsx';
    input.onchange = async () => {
      const file = input.files[0];
      if (!file) { reject(new Error('cancelled')); return; }
      const buf = await file.arrayBuffer();
      const result = loadWorkbook(buf, file.name);
      state.fileHandle = null;
      state.saveMode = 'download';
      resolve(result);
    };
    input.click();
  });
}

/* ========== テンプレート発行 ========== */
function buildTemplateWorkbook(nendo) {
  const wb = XLSX.utils.book_new();
  const y = String(nendo);
  const samples = [
    { '生徒ID': y + '001', '生徒氏名': '山田太郎', 'ふりがな': 'やまだたろう', '性別': '男', '生年月日': '2013-05-01', '出身小学校': '〇〇小', '住所': '〇〇市〇〇町1-2-3', '連絡先番号': '090-0000-0000', '連絡先氏名': '山田花子', '兄弟姉妹': '', '地区・通学': '〇〇地区・徒歩', '健康メモ': '', 'クラブ': 'サッカー', '委員会': '', '支援計画': '', '1年クラス': '1', '1年番号': '1', '在籍': STATUS_ACTIVE, '備考': 'サンプル行です。削除して使ってください' },
    { '生徒ID': y + '002', '生徒氏名': '佐藤花子', 'ふりがな': 'さとうはなこ', '性別': '女', '生年月日': '2013-08-15', '出身小学校': '△△小', '住所': '〇〇市△△町4-5-6', '連絡先番号': '090-1111-1111', '連絡先氏名': '佐藤次郎', '兄弟姉妹': '兄が3年', '地区・通学': '△△地区・自転車', '健康メモ': '', 'クラブ': '吹奏楽', '委員会': '図書', '支援計画': '', '1年クラス': '1', '1年番号': '2', '在籍': STATUS_ACTIVE, '備考': 'サンプル行です。削除して使ってください' }
  ];
  const meiboRows = samples.map(s => { const r = {}; MEIBO_COLS.forEach(h => r[h] = s[h] ?? ''); return r; });
  XLSX.utils.book_append_sheet(wb, rowsToSheet(meiboRows, MEIBO_COLS), SHEET_MEIBO);
  const recSample = [{ '日付': todayStr(), '生徒ID': y + '001', '氏名': '山田太郎', '種別': '面談', '内容': 'サンプル記録です。削除して使ってください。', '記録者': '担任' }];
  XLSX.utils.book_append_sheet(wb, rowsToSheet(recSample, RECORD_COLS), SHEET_RECORDS);
  XLSX.utils.book_append_sheet(wb, rowsToSheet([], RELATION_COLS), SHEET_RELATIONS);
  XLSX.utils.book_append_sheet(wb, rowsToSheet([], GRADE_COLS), SHEET_GRADES);
  return wb;
}

/* ========== 照合（貼り付けウィザード共通） ========== */
function activeStudents() {
  return state.meibo.filter(s => (s['在籍'] || STATUS_ACTIVE) === STATUS_ACTIVE);
}

function buildNameIndex(students) {
  const byName = new Map();
  for (const s of students) {
    const key = normName(s['生徒氏名']);
    if (!key) continue;
    if (!byName.has(key)) byName.set(key, []);
    byName.get(key).push(s);
  }
  return byName;
}

function matchByName(byName, name, kana) {
  const cands = byName.get(normName(name)) || [];
  if (cands.length === 1) return { status: 'ok', student: cands[0] };
  if (cands.length > 1) {
    if (kana) {
      const k = cands.filter(c => normName(c['ふりがな']) === normName(kana));
      if (k.length === 1) return { status: 'ok', student: k[0] };
    }
    return { status: 'dup', candidates: cands };
  }
  return { status: 'none' };
}

/* ========== 同クラス不可チェック ========== */
function checkNgPairs(classById) {
  // classById: Map(生徒ID → 新クラス)。同クラス不可ペアが同じクラスなら違反。
  const violations = [];
  for (const rel of state.relations) {
    if (String(rel['種別']).trim() !== '同クラス不可') continue;
    const a = String(rel['生徒ID_A']).trim();
    const b = String(rel['生徒ID_B']).trim();
    const ca = classById.get(a);
    const cb = classById.get(b);
    if (ca && cb && ca === cb) {
      violations.push({ a, b, cls: ca, memo: rel['理由・メモ'] || '' });
    }
  }
  return violations;
}

/* ========== 生徒検索ユーティリティ ========== */
function findStudent(id) {
  return state.meibo.find(s => String(s['生徒ID']) === String(id));
}

function studentLabel(s) {
  const g = gradeOf(s);
  const pos = (g >= 1 && g <= 3 && currentClass(s)) ? `${g}年${currentClass(s)}組${currentNumber(s)}番` : '';
  return `${s['生徒氏名']}（${s['生徒ID']}${pos ? '・' + pos : ''}）`;
}
