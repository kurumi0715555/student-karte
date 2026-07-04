/* 生徒カルテ UI */
'use strict';

/* ========== 要素キャッシュ ========== */
const $ = id => document.getElementById(id);
const els = {
  startView: $('startView'), mainView: $('mainView'), noticeArea: $('noticeArea'),
  homeView: $('homeView'), homeSearch: $('homeSearch'), homeResults: $('homeResults'), homeStats: $('homeStats'),
  nendoSelect: $('nendoSelect'), saveStatus: $('saveStatus'),
  filterGrade: $('filterGrade'), filterClass: $('filterClass'), filterStatus: $('filterStatus'),
  searchBox: $('searchBox'), toggleSensitive: $('toggleSensitive'),
  rosterTable: $('rosterTable'), emptyMessage: $('emptyMessage'),
  kartePanel: $('kartePanel'), karteName: $('karteName'), karteMeta: $('karteMeta'), karteBody: $('karteBody'),
  modalOverlay: $('modalOverlay'), modalTitle: $('modalTitle'), modalBody: $('modalBody'),
  toast: $('toast'), btnManualSave: $('btnManualSave')
};

const ui = { sortKey: 'class', sortAsc: true, karteId: null, karteTab: 'basic' };

/* ========== 初期化 ========== */
function init() {
  state.nendo = currentNendo();
  const sel = els.nendoSelect;
  for (let y = state.nendo - 4; y <= state.nendo + 1; y++) {
    const opt = document.createElement('option');
    opt.value = y; opt.textContent = `R${y}`;
    if (y === state.nendo) opt.selected = true;
    sel.appendChild(opt);
  }
  sel.addEventListener('change', () => { state.nendo = parseInt(sel.value, 10); renderAll(); });

  $('startModeNote').textContent = fsaSupported
    ? '✅ このブラウザは自動上書き保存に対応しています。編集はそのままExcelファイルに保存されます。'
    : 'ℹ このブラウザでは自動上書き保存が使えません。保存は「ダウンロードして差し替え」方式になります（Chrome / Edge 推奨）。';

  $('btnTemplate').addEventListener('click', () => {
    downloadBlob(workbookToBlob(buildTemplateWorkbook(state.nendo)), FILE_DEFAULT_NAME);
    showToast('テンプレートをダウンロードしました。Excelで入力してから開いてください。');
  });
  $('btnOpenFile').addEventListener('click', openFileFlow);

  els.filterGrade.addEventListener('change', () => { renderClassFilter(); renderTable(); });
  els.filterClass.addEventListener('change', renderTable);
  els.filterStatus.addEventListener('change', renderTable);
  els.searchBox.addEventListener('input', renderTable);
  els.toggleSensitive.addEventListener('change', renderTable);

  $('btnHome').addEventListener('click', showHome);
  $('btnHomeAdd').addEventListener('click', openAddStudentModal);
  $('btnHomeList').addEventListener('click', showList);
  $('btnHomePromote').addEventListener('click', openPromoteWizard);
  $('btnHomeGrades').addEventListener('click', openGradesImportWizard);
  $('btnHomeSummary').addEventListener('click', openSummaryModal);
  $('btnHomeBackup').addEventListener('click', () => { downloadBackup(); showToast('バックアップをダウンロードしました。'); });
  $('btnHomeManualSave').addEventListener('click', () => { manualDownloadSave(); showToast('ダウンロードしました。元のファイルと差し替えてください。'); });
  els.homeSearch.addEventListener('input', renderHomeResults);

  $('btnAddStudent').addEventListener('click', openAddStudentModal);
  $('btnPromote').addEventListener('click', openPromoteWizard);
  $('btnGradesImport').addEventListener('click', openGradesImportWizard);
  $('btnCopyTsv').addEventListener('click', openTsvCopyModal);
  $('btnSummary').addEventListener('click', openSummaryModal);
  $('btnGraduate').addEventListener('click', openGraduateModal);
  $('btnBackup').addEventListener('click', () => { downloadBackup(); showToast('バックアップをダウンロードしました。'); });
  els.btnManualSave.addEventListener('click', () => { manualDownloadSave(); showToast('ダウンロードしました。元のファイルと差し替えてください。'); });

  $('btnCloseKarte').addEventListener('click', closeKarte);
  $('btnPrintKarte').addEventListener('click', openPrintModal);
  els.kartePanel.addEventListener('click', e => { if (e.target === els.kartePanel) closeKarte(); });
  document.querySelectorAll('.karte-tab').forEach(btn => {
    btn.addEventListener('click', () => { ui.karteTab = btn.dataset.tab; renderKarte(); });
  });

  $('modalClose').addEventListener('click', closeModal);
  els.modalOverlay.addEventListener('click', e => { if (e.target === els.modalOverlay) closeModal(); });

  window.addEventListener('beforeunload', e => {
    if (state.loaded && state.dirty > 0) { e.preventDefault(); e.returnValue = ''; }
  });
}

async function openFileFlow() {
  try {
    const result = fsaSupported ? await openWithPicker() : await openWithInput();
    if (result.problems.errors.length > 0) {
      alertNotice('error', '読み込めませんでした', result.problems.errors);
      return;
    }
    els.startView.hidden = true;
    els.homeView.hidden = false;
    els.mainView.hidden = true;
    els.noticeArea.innerHTML = '';
    if (result.problems.warnings.length > 0) {
      alertNotice('warning', 'データに注意点があります', result.problems.warnings);
    }
    addBackupNotice();
    renderAll();
    onSaveStateChanged();
  } catch (e) {
    if (e && (e.name === 'AbortError' || e.message === 'cancelled')) return;
    console.error(e);
    alertNotice('error', 'ファイルを開けませんでした', [String(e.message || e)]);
  }
}

function alertNotice(kind, title, items) {
  const div = document.createElement('div');
  div.className = `notice notice-${kind}`;
  div.innerHTML = `<div><strong>${escapeHtml(title)}</strong><ul>${items.map(i => `<li>${escapeHtml(i)}</li>`).join('')}</ul></div>
    <button class="btn btn-ghost btn-small">✕</button>`;
  div.querySelector('button').addEventListener('click', () => div.remove());
  els.noticeArea.appendChild(div);
}

function addBackupNotice() {
  const div = document.createElement('div');
  div.className = 'notice notice-info';
  div.innerHTML = `<div>📦 作業を始める前に、今日のバックアップを保存しておくと安心です。</div>
    <span><button class="btn btn-primary btn-small" id="nbBackup">バックアップを保存</button>
    <button class="btn btn-ghost btn-small" id="nbSkip">今回はしない</button></span>`;
  div.querySelector('#nbBackup').addEventListener('click', () => { downloadBackup(); div.remove(); showToast('バックアップをダウンロードしました。'); });
  div.querySelector('#nbSkip').addEventListener('click', () => div.remove());
  els.noticeArea.appendChild(div);
}

/* ========== 保存ステータス ========== */
function onSaveStateChanged() {
  const st = els.saveStatus;
  st.hidden = !state.loaded;
  if (!state.loaded) return;
  if (state.saveError) {
    st.className = 'save-status error';
    st.innerHTML = '⚠ 保存に失敗（Excelで開いていませんか？） <a href="#" id="retryLink" style="color:inherit">再保存</a>';
    st.querySelector('#retryLink').addEventListener('click', e => { e.preventDefault(); retrySave(); });
  } else if (state.saveMode === 'fsa') {
    st.className = 'save-status saved';
    st.textContent = state.dirty === 0 ? '✓ 保存済み' : '保存中…';
  } else {
    st.className = state.dirty > 0 ? 'save-status dirty' : 'save-status saved';
    st.textContent = state.dirty > 0 ? `未保存の変更 ${state.dirty}件 — 「保存」を押してください` : '✓ 保存済み';
  }
  document.querySelectorAll('.manual-save-btn').forEach(b => { b.hidden = (state.saveMode === 'fsa'); });
}

/* ========== ビュー切替（ホーム ⇄ 名簿一覧） ========== */
function showHome() {
  els.mainView.hidden = true;
  els.homeView.hidden = false;
  renderHome();
  els.homeSearch.focus();
}

function showList() {
  els.homeView.hidden = true;
  els.mainView.hidden = false;
  renderClassFilter();
  renderTable();
  onSaveStateChanged();
}

function renderHome() {
  const active = activeStudents().filter(s => { const g = gradeOf(s); return g >= 1 && g <= 3; });
  const todayRecs = state.records.filter(r => String(r['日付']) === todayStr()).length;
  els.homeStats.textContent = `R${state.nendo}年度 在籍 ${active.length}名 ／ 今日書いた記録 ${todayRecs}件`;
  renderHomeResults();
  onSaveStateChanged();
}

function renderHomeResults() {
  const q = normName(els.homeSearch.value).toLowerCase();
  if (!q) { els.homeResults.innerHTML = ''; return; }
  const hits = state.meibo.filter(s => {
    const hay = (normName(s['生徒氏名']) + normName(s['ふりがな']) + String(s['生徒ID'])).toLowerCase();
    return hay.includes(q);
  });
  const shown = hits.slice(0, 30);
  els.homeResults.innerHTML = shown.map(s => {
    const g = gradeOf(s);
    const pos = (g >= 1 && g <= 3 && currentClass(s)) ? `${g}年${currentClass(s)}組${currentNumber(s)}番` : (g > 3 ? '卒業済' : '');
    const status = s['在籍'] || STATUS_ACTIVE;
    return `<button class="home-result-item" data-id="${escapeHtml(String(s['生徒ID']))}">
      <span class="hr-name">${escapeHtml(s['生徒氏名'])}</span>
      <span class="hr-meta">${escapeHtml(s['ふりがな'] || '')}　${escapeHtml(pos)}　ID:${escapeHtml(String(s['生徒ID']))}</span>
      ${status !== STATUS_ACTIVE ? `<span class="hr-badge badge badge-status">${escapeHtml(status)}</span>` : ''}
    </button>`;
  }).join('') + (hits.length > 30 ? `<p class="home-result-more">ほか ${hits.length - 30} 名 — さらに絞り込んでください</p>` : '');
  els.homeResults.querySelectorAll('.home-result-item').forEach(btn => {
    btn.addEventListener('click', () => openKarte(btn.dataset.id));
  });
}

/* ========== 一覧テーブル ========== */
const BASE_COLS = [
  { key: '生徒ID', label: 'ID' }, { key: '生徒氏名', label: '氏名' }, { key: 'ふりがな', label: 'ふりがな' },
  { key: '_grade', label: '学年', computed: true }, { key: '_class', label: 'クラス', computed: true },
  { key: '_num', label: '番号', computed: true },
  { key: '性別', label: '性別' }, { key: '出身小学校', label: '出身小' },
  { key: 'クラブ', label: 'クラブ' }, { key: '委員会', label: '委員会' }
];
const SENS_VIEW_COLS = SENSITIVE_COLS.map(c => ({ key: c, label: c, sensitive: true }));
const TAIL_COLS = [{ key: '在籍', label: '在籍' }, { key: '備考', label: '備考' }];

function visibleCols() {
  return els.toggleSensitive.checked
    ? [...BASE_COLS, ...SENS_VIEW_COLS, ...TAIL_COLS]
    : [...BASE_COLS, ...TAIL_COLS];
}

function cellValue(s, key) {
  if (key === '_grade') { const g = gradeOf(s); return (g >= 1 && g <= 3) ? g + '年' : (g > 3 ? '卒業済' : '—'); }
  if (key === '_class') return currentClass(s);
  if (key === '_num') return currentNumber(s);
  return String(s[key] ?? '');
}

function filteredStudents() {
  const gsel = els.filterGrade.value;
  const csel = els.filterClass.value;
  const ssel = els.filterStatus.value;
  const q = normName(els.searchBox.value).toLowerCase();
  return state.meibo.filter(s => {
    const g = gradeOf(s);
    const status = s['在籍'] || STATUS_ACTIVE;
    if (ssel !== 'all' && status !== ssel) return false;
    if (gsel === 'other') { if (g >= 1 && g <= 3) return false; }
    else if (gsel !== 'all' && g !== parseInt(gsel, 10)) return false;
    if (csel !== 'all' && currentClass(s) !== csel) return false;
    if (q) {
      const hay = (normName(s['生徒氏名']) + normName(s['ふりがな']) + String(s['生徒ID'])).toLowerCase();
      if (!hay.includes(q)) return false;
    }
    return true;
  });
}

function sortStudents(list) {
  const key = ui.sortKey, asc = ui.sortAsc ? 1 : -1;
  const num = v => { const n = parseFloat(v); return isNaN(n) ? Infinity : n; };
  return list.slice().sort((a, b) => {
    let r = 0;
    if (key === 'class') {
      r = (gradeOf(a) - gradeOf(b))
        || String(currentClass(a)).localeCompare(String(currentClass(b)), 'ja', { numeric: true })
        || (num(currentNumber(a)) - num(currentNumber(b)));
    } else if (key === 'kana') {
      r = String(a['ふりがな'] || '').localeCompare(String(b['ふりがな'] || ''), 'ja');
    } else if (key === 'id') {
      r = String(a['生徒ID']).localeCompare(String(b['生徒ID']), 'ja', { numeric: true });
    } else {
      r = cellValue(a, key).localeCompare(cellValue(b, key), 'ja', { numeric: true });
    }
    return r * asc;
  });
}

function renderClassFilter() {
  const gsel = els.filterGrade.value;
  const classes = new Set();
  for (const s of state.meibo) {
    const g = gradeOf(s);
    if (gsel !== 'all' && gsel !== 'other' && g !== parseInt(gsel, 10)) continue;
    const c = currentClass(s);
    if (c) classes.add(c);
  }
  const cur = els.filterClass.value;
  els.filterClass.innerHTML = '<option value="all">全クラス</option>' +
    [...classes].sort((a, b) => a.localeCompare(b, 'ja', { numeric: true })).map(c => `<option value="${escapeHtml(c)}">${escapeHtml(c)}組</option>`).join('');
  if ([...els.filterClass.options].some(o => o.value === cur)) els.filterClass.value = cur;
}

function renderTable() {
  const cols = visibleCols();
  const rows = sortStudents(filteredStudents());
  const sortMark = k => (ui.sortKey === k || (k === '_class' && ui.sortKey === 'class')) ? `<span class="sort-arrow">${ui.sortAsc ? '▲' : '▼'}</span>` : '';
  let html = '<thead><tr>' + cols.map(c =>
    `<th data-key="${c.key}" class="${c.sensitive ? 'sensitive-col' : ''}">${escapeHtml(c.label)}${sortMark(c.key)}</th>`
  ).join('') + '</tr></thead><tbody>';
  for (const s of rows) {
    const status = s['在籍'] || STATUS_ACTIVE;
    html += `<tr data-id="${escapeHtml(String(s['生徒ID']))}" class="${status !== STATUS_ACTIVE ? 'row-inactive' : ''}">`;
    for (const c of cols) {
      let v = escapeHtml(cellValue(s, c.key));
      if (c.key === '支援計画' && v === 'あり') v = '<span class="badge badge-support">支援あり</span>';
      if (c.key === '在籍' && status !== STATUS_ACTIVE) v = `<span class="badge badge-status">${escapeHtml(status)}</span>`;
      html += `<td class="${c.sensitive ? 'sensitive-col' : ''}">${v}</td>`;
    }
    html += '</tr>';
  }
  html += '</tbody>';
  els.rosterTable.innerHTML = html;
  els.emptyMessage.hidden = rows.length > 0;

  els.rosterTable.querySelectorAll('th').forEach(th => {
    th.addEventListener('click', () => {
      const k = th.dataset.key === '_class' ? 'class' : (th.dataset.key === 'ふりがな' ? 'kana' : (th.dataset.key === '生徒ID' ? 'id' : th.dataset.key));
      if (ui.sortKey === k) ui.sortAsc = !ui.sortAsc; else { ui.sortKey = k; ui.sortAsc = true; }
      renderTable();
    });
  });
  els.rosterTable.querySelectorAll('tbody tr').forEach(tr => {
    tr.addEventListener('click', () => openKarte(tr.dataset.id));
  });
}

function renderAll() {
  if (!state.loaded) return;
  if (!els.homeView.hidden) renderHome();
  if (!els.mainView.hidden) { renderClassFilter(); renderTable(); }
  onSaveStateChanged();
}

/* ========== 生徒カルテ ========== */
function openKarte(id) {
  ui.karteId = id;
  ui.karteTab = 'basic';
  document.querySelectorAll('.karte-tab').forEach(b => b.classList.toggle('active', b.dataset.tab === 'basic'));
  els.kartePanel.hidden = false;
  renderKarte();
}
function closeKarte() { els.kartePanel.hidden = true; ui.karteId = null; }

function renderKarte() {
  const s = findStudent(ui.karteId);
  if (!s) { closeKarte(); return; }
  els.karteName.textContent = `${s['生徒氏名']}（${s['ふりがな'] || ''}）`;
  const g = gradeOf(s);
  els.karteMeta.textContent = `ID: ${s['生徒ID']} ／ ` +
    ((g >= 1 && g <= 3) ? `R${state.nendo}年度 ${g}年${currentClass(s) || '?'}組${currentNumber(s) || '?'}番` : '在学対象外') +
    ` ／ ${s['在籍'] || STATUS_ACTIVE}`;
  document.querySelectorAll('.karte-tab').forEach(b => b.classList.toggle('active', b.dataset.tab === ui.karteTab));
  if (ui.karteTab === 'basic') renderKarteBasic(s);
  else if (ui.karteTab === 'records') renderKarteRecords(s);
  else if (ui.karteTab === 'grades') renderKarteGrades(s);
  else renderKarteRelations(s);
}

function renderKarteBasic(s) {
  const wideCols = ['住所', '健康メモ', '備考', '兄弟姉妹'];
  let html = '<div class="field-grid">';
  for (const h of state.meiboHeaders) {
    const sens = SENSITIVE_COLS.includes(h);
    const wide = wideCols.includes(h);
    const val = escapeHtml(String(s[h] ?? ''));
    const hAttr = escapeHtml(h);
    let input;
    if (h === '生徒ID') input = `<input type="text" data-col="${hAttr}" value="${val}" readonly>`;
    else if (h === '在籍') input = `<select data-col="${hAttr}">${[STATUS_ACTIVE, '転出', '卒業'].map(o => `<option ${o === (s[h] || STATUS_ACTIVE) ? 'selected' : ''}>${o}</option>`).join('')}</select>`;
    else if (h === '支援計画') input = `<select data-col="${hAttr}"><option value="" ${!s[h] ? 'selected' : ''}>—</option><option ${s[h] === 'あり' ? 'selected' : ''}>あり</option><option ${s[h] === 'なし' ? 'selected' : ''}>なし</option></select>`;
    else input = `<input type="text" data-col="${hAttr}" value="${val}">`;
    html += `<div class="field ${wide ? 'wide' : ''} ${sens ? 'sensitive' : ''}" data-sensitive="${sens}">
      <label>${sens ? '<span class="sens-mark">🔒</span> ' : ''}${escapeHtml(h)}</label>${input}</div>`;
  }
  html += '</div><div class="karte-actions"><button class="btn btn-primary" id="karteSaveBasic">この内容で保存</button></div>';
  els.karteBody.innerHTML = html;
  $('karteSaveBasic').addEventListener('click', async () => {
    els.karteBody.querySelectorAll('[data-col]').forEach(inp => {
      if (inp.dataset.col !== '生徒ID') s[inp.dataset.col] = inp.value.trim();
    });
    await persist();
    showToast('保存しました。');
    renderKarte();
    renderAll();
  });
}

function renderKarteRecords(s) {
  const id = String(s['生徒ID']);
  const list = state.records
    .map((r, idx) => ({ r, idx }))
    .filter(x => String(x.r['生徒ID']) === id)
    .sort((a, b) => String(b.r['日付']).localeCompare(String(a.r['日付'])));
  let html = `<div class="add-form">
    <div class="form-row">
      <input type="date" id="recDate" value="${todayStr()}">
      <select id="recType">${RECORD_TYPES.map(t => `<option>${t}</option>`).join('')}</select>
      <input type="text" id="recBy" placeholder="記録者" style="width:110px">
    </div>
    <textarea id="recBody" placeholder="内容を入力"></textarea>
    <div class="form-row" style="justify-content:flex-end"><button class="btn btn-primary btn-small" id="recAdd">＋ 記録を追加</button></div>
  </div>`;
  if (list.length === 0) html += '<p style="color:var(--gray-dark)">記録はまだありません。</p>';
  for (const { r, idx } of list) {
    html += `<div class="record-item">
      <div class="record-head"><span>${escapeHtml(r['日付'])}</span><span class="record-type">${escapeHtml(r['種別'])}</span>
      <span>${escapeHtml(r['記録者'] || '')}</span>
      <button class="btn btn-ghost btn-small" data-del="${idx}" style="margin-left:auto">削除</button></div>
      <p>${escapeHtml(r['内容'])}</p></div>`;
  }
  els.karteBody.innerHTML = html;
  $('recAdd').addEventListener('click', async () => {
    const body = $('recBody').value.trim();
    if (!body) { showToast('内容を入力してください。', true); return; }
    state.records.push({ '日付': $('recDate').value || todayStr(), '生徒ID': id, '氏名': s['生徒氏名'], '種別': $('recType').value, '内容': body, '記録者': $('recBy').value.trim() });
    await persist();
    showToast('記録を追加しました。');
    renderKarte();
  });
  els.karteBody.querySelectorAll('[data-del]').forEach(btn => {
    btn.addEventListener('click', () => {
      openConfirm('記録の削除', 'この記録を削除します。元に戻せません。よろしいですか？', async () => {
        state.records.splice(parseInt(btn.dataset.del, 10), 1);
        await persist();
        showToast('削除しました。');
        renderKarte();
      });
    });
  });
}

function renderKarteGrades(s) {
  const id = String(s['生徒ID']);
  const list = state.grades.filter(r => String(r['生徒ID']) === id);
  const groups = new Map();
  for (const r of list) {
    const key = `R${r['年度']}・${r['学期']}・${r['テスト名']}`;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(r);
  }
  let html = `<div class="add-form"><div class="form-row">
    <input type="number" id="grNendo" value="${state.nendo}" style="width:80px" title="年度(令和)">
    <select id="grTerm"><option>1学期</option><option>2学期</option><option>3学期</option><option>その他</option></select>
    <input type="text" id="grTest" placeholder="テスト名" style="width:140px">
    <input type="text" id="grSubject" placeholder="教科" style="width:90px">
    <input type="number" id="grScore" placeholder="点数" style="width:80px">
    <button class="btn btn-primary btn-small" id="grAdd">＋ 追加</button>
  </div></div>`;
  if (groups.size === 0) html += '<p style="color:var(--gray-dark)">成績はまだありません。クラス全体の一括登録はツールバーの「成績取り込み」が便利です。</p>';
  for (const [key, rows] of [...groups.entries()].sort((a, b) => b[0].localeCompare(a[0], 'ja', { numeric: true }))) {
    html += `<div class="summary-block"><h3>${escapeHtml(key)}</h3><table class="summary-table"><tr>` +
      rows.map(r => `<th>${escapeHtml(r['教科'])}</th>`).join('') + '</tr><tr>' +
      rows.map(r => `<td>${escapeHtml(r['点数'])}</td>`).join('') + '</tr></table></div>';
  }
  els.karteBody.innerHTML = html;
  $('grAdd').addEventListener('click', async () => {
    const test = $('grTest').value.trim(), subj = $('grSubject').value.trim(), score = $('grScore').value.trim();
    if (!test || !subj || score === '') { showToast('テスト名・教科・点数を入力してください。', true); return; }
    state.grades.push({ '生徒ID': id, '氏名': s['生徒氏名'], '年度': $('grNendo').value, '学期': $('grTerm').value, 'テスト名': test, '教科': subj, '点数': score });
    await persist();
    showToast('成績を追加しました。');
    renderKarte();
  });
}

function renderKarteRelations(s) {
  const id = String(s['生徒ID']);
  const list = state.relations
    .map((r, idx) => ({ r, idx }))
    .filter(x => String(x.r['生徒ID_A']) === id || String(x.r['生徒ID_B']) === id);
  const others = activeStudents().filter(o => String(o['生徒ID']) !== id);
  let html = `<div class="add-form"><div class="form-row">
    <select id="relType">${RELATION_TYPES.map(t => `<option>${t}</option>`).join('')}</select>
    <select id="relOther" style="flex:1;min-width:200px">
      <option value="">相手の生徒を選択…</option>
      ${others.map(o => `<option value="${escapeHtml(String(o['生徒ID']))}">${escapeHtml(studentLabel(o))}</option>`).join('')}
    </select></div>
    <input type="text" id="relMemo" placeholder="理由・メモ（任意）">
    <div class="form-row" style="justify-content:flex-end"><button class="btn btn-primary btn-small" id="relAdd">＋ 関係を追加</button></div>
  </div>
  <p class="hint" style="background:var(--info-bg);border-radius:6px;padding:8px 12px;font-size:0.85em">
  「同クラス不可」の関係は、進級取り込みのときに自動でチェックされます。</p>`;
  if (list.length === 0) html += '<p style="color:var(--gray-dark)">登録された関係はありません。</p>';
  for (const { r, idx } of list) {
    const otherId = String(r['生徒ID_A']) === id ? String(r['生徒ID_B']) : String(r['生徒ID_A']);
    const other = findStudent(otherId);
    html += `<div class="record-item">
      <div class="record-head"><span class="record-type">${escapeHtml(r['種別'])}</span>
      <span>${other ? escapeHtml(studentLabel(other)) : 'ID:' + escapeHtml(otherId) + '（名簿に見つかりません）'}</span>
      <button class="btn btn-ghost btn-small" data-del="${idx}" style="margin-left:auto">削除</button></div>
      ${r['理由・メモ'] ? `<p>${escapeHtml(r['理由・メモ'])}</p>` : ''}</div>`;
  }
  els.karteBody.innerHTML = html;
  $('relAdd').addEventListener('click', async () => {
    const otherId = $('relOther').value;
    if (!otherId) { showToast('相手の生徒を選択してください。', true); return; }
    const other = findStudent(otherId);
    state.relations.push({ '生徒ID_A': id, '氏名A': s['生徒氏名'], '生徒ID_B': otherId, '氏名B': other ? other['生徒氏名'] : '', '種別': $('relType').value, '理由・メモ': $('relMemo').value.trim() });
    await persist();
    showToast('関係を追加しました。');
    renderKarte();
  });
  els.karteBody.querySelectorAll('[data-del]').forEach(btn => {
    btn.addEventListener('click', () => {
      openConfirm('関係の削除', 'この関係を削除します。よろしいですか？', async () => {
        state.relations.splice(parseInt(btn.dataset.del, 10), 1);
        await persist();
        showToast('削除しました。');
        renderKarte();
      });
    });
  });
}

/* ========== モーダル基盤 ========== */
function openModal(title, bodyHtml) {
  els.modalTitle.textContent = title;
  els.modalBody.innerHTML = bodyHtml;
  els.modalOverlay.hidden = false;
}
function closeModal() { els.modalOverlay.hidden = true; els.modalBody.innerHTML = ''; }

function openConfirm(title, message, onOk) {
  openModal(title, `<p>${escapeHtml(message)}</p>
    <div class="modal-actions">
      <button class="btn" id="cfCancel">キャンセル</button>
      <button class="btn btn-danger" id="cfOk">実行する</button>
    </div>`);
  $('cfCancel').addEventListener('click', closeModal);
  $('cfOk').addEventListener('click', () => { closeModal(); onOk(); });
}

/* ========== 列マッピングUI（貼り付けウィザード共通） ========== */
function guessHasHeader(rows) {
  if (rows.length < 2) return false;
  return rows[0].some(c => /氏名|名前|ふりがな|クラス|組|番号|出席/.test(c));
}

function mappingSelect(id, rows, guessPatterns, allowNone) {
  const sample = rows[0] || [];
  const opts = sample.map((c, i) =>
    `<option value="${i}">${i + 1}列目（${escapeHtml(String(c).slice(0, 12)) || '空'}）</option>`).join('');
  let guessed = -1;
  if (guessPatterns) {
    sample.forEach((c, i) => { if (guessed < 0 && guessPatterns.test(c)) guessed = i; });
  }
  return { html: `<select id="${id}">${allowNone ? '<option value="-1">（使わない）</option>' : ''}${opts}</select>`, guessed };
}

/* ========== 進級取り込みウィザード ========== */
function openPromoteWizard() {
  openModal('進級取り込み — 新しいクラス・番号を一括登録', `
    <div class="hint">校務支援システム等の<strong>新クラス編成表</strong>（新クラス・新番号・氏名の一覧）をExcelでコピーし、下の欄に貼り付けてください。氏名で名簿と自動照合します。<br>
    💡 ヘッダーの年度を<strong>進級後の年度</strong>（現在: R${state.nendo}）に切り替えてから取り込むと、転出漏れの検出が正確になります。</div>
    <div class="form-row" style="display:flex;gap:12px;align-items:center;margin:10px 0">
      <label>書き込み先：
        <select id="pwGrade">
          <option value="1">1年クラス・番号</option>
          <option value="2" selected>2年クラス・番号</option>
          <option value="3">3年クラス・番号</option>
        </select>
      </label>
    </div>
    <textarea class="paste-area" id="pwPaste" placeholder="ここにExcelからコピーした表を貼り付け（タブ区切り）"></textarea>
    <div class="modal-actions"><button class="btn btn-primary" id="pwParse">照合する →</button></div>
    <div id="pwResult"></div>`);
  $('pwParse').addEventListener('click', promoteParse);
}

function promoteParse() {
  const rows = parseTSV($('pwPaste').value);
  if (rows.length === 0) { showToast('貼り付けデータがありません。', true); return; }
  const targetGrade = parseInt($('pwGrade').value, 10);
  const hasHeader = guessHasHeader(rows);

  const mCls = mappingSelect('pwColCls', rows, /クラス|組/, false);
  const mNum = mappingSelect('pwColNum', rows, /番号|出席/, false);
  const mName = mappingSelect('pwColName', rows, /氏名|名前/, false);
  const mKana = mappingSelect('pwColKana', rows, /ふりがな|かな/, true);
  $('pwResult').innerHTML = `
    <div class="hint">列の割り当てを確認してください（1行目のデータから推測しています）。</div>
    <div style="display:flex;gap:14px;flex-wrap:wrap;margin:8px 0">
      <label>新クラス ${mCls.html}</label>
      <label>新番号 ${mNum.html}</label>
      <label>氏名 ${mName.html}</label>
      <label>ふりがな ${mKana.html}</label>
      <label><input type="checkbox" id="pwHeader" ${hasHeader ? 'checked' : ''}> 1行目は見出し</label>
    </div>
    <div class="modal-actions"><button class="btn btn-primary" id="pwMatch">この割り当てで照合 →</button></div>
    <div id="pwPreview"></div>`;
  if (mCls.guessed >= 0) $('pwColCls').value = mCls.guessed;
  if (mNum.guessed >= 0) $('pwColNum').value = mNum.guessed;
  if (mName.guessed >= 0) $('pwColName').value = mName.guessed;
  $('pwColKana').value = mKana.guessed >= 0 ? mKana.guessed : -1;
  $('pwMatch').addEventListener('click', () => promoteMatch(rows, targetGrade));
}

function promoteMatch(allRows, targetGrade) {
  const ci = {
    cls: parseInt($('pwColCls').value, 10),
    num: parseInt($('pwColNum').value, 10),
    name: parseInt($('pwColName').value, 10),
    kana: parseInt($('pwColKana').value, 10)
  };
  const rows = $('pwHeader').checked ? allRows.slice(1) : allRows;
  const byName = buildNameIndex(activeStudents());
  const resolved = rows.map(cells => {
    const name = cells[ci.name] || '';
    const kana = ci.kana >= 0 ? cells[ci.kana] : '';
    const m = matchByName(byName, name, kana);
    return { name, kana, cls: cells[ci.cls] || '', num: cells[ci.num] || '', match: m, chosenId: m.status === 'ok' ? String(m.student['生徒ID']) : '' };
  });

  const allActive = activeStudents();
  const matchedIds = new Set(resolved.filter(r => r.chosenId).map(r => r.chosenId));
  const options = allActive.map(o => `<option value="${escapeHtml(String(o['生徒ID']))}">${escapeHtml(studentLabel(o))}</option>`).join('');

  let html = `<table class="preview-table"><thead><tr><th>状態</th><th>貼り付けた氏名</th><th>新クラス</th><th>新番号</th><th>名簿との対応</th></tr></thead><tbody>`;
  resolved.forEach((r, i) => {
    let cls = 'match-ok', label = '';
    if (r.match.status === 'ok') label = escapeHtml(studentLabel(r.match.student));
    else {
      cls = 'match-ng';
      label = `<select data-fix="${i}"><option value="">— 対応する生徒を選択（or 無視） —</option>${options}</select>`;
    }
    html += `<tr class="${cls}"><td>${r.match.status === 'ok' ? '✓' : (r.match.status === 'dup' ? '同名複数' : '見つからない')}</td>
      <td>${escapeHtml(r.name)}</td><td>${escapeHtml(r.cls)}</td><td>${escapeHtml(r.num)}</td><td>${label}</td></tr>`;
  });
  html += '</tbody></table>';

  const leftover = allActive.filter(s => {
    const g = gradeOf(s);
    return g === targetGrade && !matchedIds.has(String(s['生徒ID']));
  });
  if (leftover.length > 0) {
    html += `<div class="notice notice-warning" style="margin-top:10px"><div><strong>編成表に見つからなかった在籍生徒（転出漏れ・貼り忘れの可能性）:</strong>
      <ul>${leftover.map(s => `<li>${escapeHtml(studentLabel(s))}</li>`).join('')}</ul>
      <span style="font-size:0.85em">※ 表示中の R${state.nendo} 年度で ${targetGrade} 年生になる在籍生徒を基準に判定しています。</span></div></div>`;
  }
  html += `<div id="pwNgArea"></div>
    <div class="modal-actions"><button class="btn btn-primary" id="pwApply">確定して書き込む</button></div>`;
  $('pwPreview').innerHTML = html;

  $('pwPreview').querySelectorAll('[data-fix]').forEach(sel => {
    sel.addEventListener('change', () => { resolved[parseInt(sel.dataset.fix, 10)].chosenId = sel.value; });
  });

  $('pwApply').addEventListener('click', async () => {
    const assign = resolved.filter(r => r.chosenId);
    // 同一生徒への重複割当を検知（後の行が無言で上書きするのを防ぐ）
    const idCounts = new Map();
    for (const r of assign) idCounts.set(r.chosenId, (idCounts.get(r.chosenId) || 0) + 1);
    const dups = [...idCounts.entries()].filter(([, c]) => c > 1);
    if (dups.length > 0 && !$('pwNgArea').dataset.dupAck) {
      $('pwNgArea').dataset.dupAck = '1';
      $('pwNgArea').innerHTML = `<div class="notice notice-error"><div>
        <strong>⚠ 同じ生徒に複数の行が割り当てられています（最後の行の値で上書きされます）:</strong>
        <ul>${dups.map(([id]) => { const s = findStudent(id); return `<li>${escapeHtml(s ? s['生徒氏名'] : id)}（ID: ${escapeHtml(id)}）</li>`; }).join('')}</ul>
        対応表を見直すか、このまま進める場合はもう一度「確定して書き込む」を押してください。</div></div>`;
      return;
    }
    const classById = new Map(assign.map(r => [r.chosenId, String(r.cls)]));
    const violations = checkNgPairs(classById);
    if (violations.length > 0 && !$('pwNgArea').dataset.acknowledged) {
      $('pwNgArea').dataset.acknowledged = '1';
      $('pwNgArea').innerHTML = `<div class="notice notice-error"><div>
        <strong>⚠ 「同クラス不可」のペアが同じクラスになっています:</strong>
        <ul>${violations.map(v => {
          const sa = findStudent(v.a), sb = findStudent(v.b);
          return `<li>${escapeHtml(sa ? sa['生徒氏名'] : v.a)} と ${escapeHtml(sb ? sb['生徒氏名'] : v.b)} → ${escapeHtml(v.cls)}組${v.memo ? '（' + escapeHtml(v.memo) + '）' : ''}</li>`;
        }).join('')}</ul>
        このまま書き込む場合は、もう一度「確定して書き込む」を押してください。</div></div>`;
      return;
    }
    for (const r of assign) {
      const s = findStudent(r.chosenId);
      if (!s) continue;
      s[classColOf(targetGrade)] = String(r.cls);
      s[numColOf(targetGrade)] = String(r.num);
    }
    await persist();
    closeModal();
    showToast(`${assign.length}名のクラス・番号を書き込みました。`);
    renderAll();
  });
}

/* ========== 成績取り込みウィザード ========== */
function openGradesImportWizard() {
  openModal('成績取り込み — 採点表を一括登録', `
    <div class="hint">📄 <strong>おすすめ：採点用テンプレートから始める</strong><br>
    クラスを選んで発行すると、生徒ID・番号・氏名入り（番号順）のExcelができます。
    点数を入力して全体をコピーし、下の欄に貼り付ければ<strong>IDで確実に照合</strong>されます（同姓同名でも間違えません）。</div>
    <div style="display:flex;gap:10px;align-items:center;margin:8px 0;flex-wrap:wrap">
      <select id="gtGrade"><option value="1">1年</option><option value="2">2年</option><option value="3">3年</option></select>
      <select id="gtClass"></select>
      <button class="btn btn-secondary btn-small" id="gtIssue">📄 採点用テンプレートを発行</button>
    </div>
    <hr style="border:none;border-top:1px solid var(--gray-light);margin:14px 0">
    <div class="hint">手持ちの採点表を貼り付けることもできます。1行目に見出し（氏名・教科名）が必要です。<br>
    例： <code>氏名 [Tab] 国語 [Tab] 数学 [Tab] 英語 …</code>（この場合は氏名で照合します）</div>
    <div style="display:flex;gap:12px;flex-wrap:wrap;margin:10px 0">
      <label>年度(令和) <input type="number" id="giNendo" value="${state.nendo}" style="width:70px"></label>
      <label>学期 <select id="giTerm"><option>1学期</option><option>2学期</option><option>3学期</option><option>その他</option></select></label>
      <label>テスト名 <input type="text" id="giTest" placeholder="例：期末テスト" style="width:160px"></label>
    </div>
    <textarea class="paste-area" id="giPaste" placeholder="ここに採点表を貼り付け（1行目は見出し）"></textarea>
    <div class="modal-actions"><button class="btn btn-primary" id="giParse">照合する →</button></div>
    <div id="giResult"></div>`);
  $('giParse').addEventListener('click', gradesParse);
  const fillClasses = () => {
    const g = parseInt($('gtGrade').value, 10);
    const classes = [...new Set(activeStudents().filter(s => gradeOf(s) === g).map(s => currentClass(s)).filter(c => c))]
      .sort((a, b) => a.localeCompare(b, 'ja', { numeric: true }));
    $('gtClass').innerHTML = classes.length
      ? classes.map(c => `<option value="${escapeHtml(c)}">${escapeHtml(c)}組</option>`).join('')
      : '<option value="">クラスなし</option>';
  };
  $('gtGrade').addEventListener('change', fillClasses);
  fillClasses();
  $('gtIssue').addEventListener('click', downloadScoreTemplate);
}

function downloadScoreTemplate() {
  const g = parseInt($('gtGrade').value, 10);
  const cls = $('gtClass').value;
  if (!cls) { showToast('該当するクラスがありません。', true); return; }
  const num = v => { const n = parseFloat(v); return isNaN(n) ? Infinity : n; };
  const members = activeStudents()
    .filter(s => gradeOf(s) === g && currentClass(s) === cls)
    .sort((a, b) => num(currentNumber(a)) - num(currentNumber(b)));
  if (members.length === 0) { showToast('該当する生徒がいません。', true); return; }
  const SUBJECTS = ['国語', '数学', '英語', '理科', '社会'];
  const headers = ['生徒ID', '番号', '氏名', ...SUBJECTS];
  const rows = members.map(s => {
    const r = { '生徒ID': String(s['生徒ID']), '番号': currentNumber(s), '氏名': s['生徒氏名'] };
    SUBJECTS.forEach(x => r[x] = '');
    return r;
  });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, rowsToSheet(rows, headers), '採点表');
  downloadBlob(workbookToBlob(wb), `採点表_R${state.nendo}_${g}年${cls}組.xlsx`);
  showToast(`${g}年${cls}組 ${members.length}名の採点用テンプレートをダウンロードしました。`);
}

function gradesParse() {
  const test = $('giTest').value.trim();
  if (!test) { showToast('テスト名を入力してください。', true); return; }
  const rows = parseTSV($('giPaste').value);
  if (rows.length < 2) { showToast('見出し行とデータ行が必要です。', true); return; }
  const header = rows[0];
  let nameCol = header.findIndex(h => /氏名|名前/.test(h));
  if (nameCol < 0) nameCol = 0;
  const skipPattern = /氏名|名前|ふりがな|かな|クラス|組|番号|出席|ID|合計|平均|順位/i;
  const subjectCols = header.map((h, i) => ({ h, i })).filter(x => x.i !== nameCol && x.h && !skipPattern.test(x.h));
  if (subjectCols.length === 0) { showToast('教科の列が見つかりません。1行目に教科名の見出しが必要です。', true); return; }

  // 生徒ID列があればIDで照合（テンプレート発行経由。同姓同名でも確実）
  const idCol = header.findIndex(h => /生徒\s*ID|^ID$/i.test(String(h).trim()));
  const byName = buildNameIndex(activeStudents());
  const resolved = rows.slice(1).map(cells => {
    if (idCol >= 0 && String(cells[idCol] || '').trim()) {
      const st = findStudent(String(cells[idCol]).trim());
      const name = cells[nameCol] || (st ? st['生徒氏名'] : String(cells[idCol]));
      return { name, cells, match: st ? { status: 'ok', student: st } : { status: 'none' }, chosenId: st ? String(st['生徒ID']) : '' };
    }
    const name = cells[nameCol] || '';
    const m = matchByName(byName, name, '');
    return { name, cells, match: m, chosenId: m.status === 'ok' ? String(m.student['生徒ID']) : '' };
  });
  const options = activeStudents().map(o => `<option value="${escapeHtml(String(o['生徒ID']))}">${escapeHtml(studentLabel(o))}</option>`).join('');

  let html = `<p style="margin:10px 0 0">認識した教科: <strong>${subjectCols.map(x => escapeHtml(x.h)).join('・')}</strong>
    ${idCol >= 0 ? '　<span style="color:var(--success)">✓ 生徒ID列で照合します</span>' : ''}</p>
    <table class="preview-table"><thead><tr><th>状態</th><th>氏名</th>${subjectCols.map(x => `<th>${escapeHtml(x.h)}</th>`).join('')}<th>対応</th></tr></thead><tbody>`;
  resolved.forEach((r, i) => {
    const ok = r.match.status === 'ok';
    html += `<tr class="${ok ? 'match-ok' : 'match-ng'}"><td>${ok ? '✓' : '要確認'}</td><td>${escapeHtml(r.name)}</td>` +
      subjectCols.map(x => `<td>${escapeHtml(r.cells[x.i] || '')}</td>`).join('') +
      `<td>${ok ? escapeHtml(studentLabel(r.match.student)) : `<select data-fix="${i}"><option value="">— 選択 or 無視 —</option>${options}</select>`}</td></tr>`;
  });
  html += `</tbody></table><div id="giWarnArea"></div>
    <div class="modal-actions"><button class="btn btn-primary" id="giApply">確定して登録</button></div>`;
  $('giResult').innerHTML = html;

  $('giResult').querySelectorAll('[data-fix]').forEach(sel => {
    sel.addEventListener('change', () => { resolved[parseInt(sel.dataset.fix, 10)].chosenId = sel.value; });
  });

  $('giApply').addEventListener('click', async () => {
    // 同一生徒への重複割当を検知
    const idCounts = new Map();
    for (const r of resolved) if (r.chosenId) idCounts.set(r.chosenId, (idCounts.get(r.chosenId) || 0) + 1);
    const dups = [...idCounts.entries()].filter(([, c]) => c > 1);
    if (dups.length > 0 && !$('giWarnArea').dataset.dupAck) {
      $('giWarnArea').dataset.dupAck = '1';
      $('giWarnArea').innerHTML = `<div class="notice notice-error"><div>
        <strong>⚠ 同じ生徒に複数の行が割り当てられています（成績が重複登録されます）:</strong>
        <ul>${dups.map(([id]) => { const s = findStudent(id); return `<li>${escapeHtml(s ? s['生徒氏名'] : id)}（ID: ${escapeHtml(id)}）</li>`; }).join('')}</ul>
        対応表を見直すか、このまま進める場合はもう一度「確定して登録」を押してください。</div></div>`;
      return;
    }
    let count = 0;
    for (const r of resolved) {
      if (!r.chosenId) continue;
      const s = findStudent(r.chosenId);
      for (const x of subjectCols) {
        const score = (r.cells[x.i] || '').trim();
        if (score === '') continue;
        state.grades.push({ '生徒ID': r.chosenId, '氏名': s ? s['生徒氏名'] : '', '年度': $('giNendo').value, '学期': $('giTerm').value, 'テスト名': test, '教科': x.h, '点数': score });
        count++;
      }
    }
    await persist();
    closeModal();
    showToast(`${count}件の成績を登録しました。`);
  });
}

/* ========== 生徒追加（転入・新入生一括） ========== */
function openAddStudentModal() {
  openModal('生徒を追加', `
    <div class="hint">転入生など1名の追加はこちら。新入学の一括登録は、テンプレートのExcelに直接貼り付けるのが簡単です（氏名・ふりがな・クラス・番号などを列ごと貼り付け → IDはオートフィルで連番）。</div>
    <div class="field-grid" style="margin-top:12px">
      <div class="field"><label>氏名（必須）</label><input type="text" id="asName"></div>
      <div class="field"><label>ふりがな</label><input type="text" id="asKana"></div>
      <div class="field"><label>学年</label><select id="asGrade"><option value="1">1年</option><option value="2">2年</option><option value="3">3年</option></select></div>
      <div class="field"><label>クラス</label><input type="text" id="asClass" placeholder="例：3"></div>
      <div class="field"><label>番号</label><input type="text" id="asNum" placeholder="例：31"></div>
      <div class="field"><label>生徒ID（自動発番）</label><input type="text" id="asId" readonly></div>
    </div>
    <div class="modal-actions"><button class="btn btn-primary" id="asAdd">追加する</button></div>`);
  const updateId = () => {
    const g = parseInt($('asGrade').value, 10);
    $('asId').value = nextIdFor(state.nendo - g + 1);
  };
  $('asGrade').addEventListener('change', updateId);
  updateId();
  $('asAdd').addEventListener('click', async () => {
    const name = $('asName').value.trim();
    if (!name) { showToast('氏名を入力してください。', true); return; }
    const g = parseInt($('asGrade').value, 10);
    const row = {};
    state.meiboHeaders.forEach(h => row[h] = '');
    row['生徒ID'] = $('asId').value;
    row['生徒氏名'] = name;
    row['ふりがな'] = $('asKana').value.trim();
    row[classColOf(g)] = $('asClass').value.trim();
    row[numColOf(g)] = $('asNum').value.trim();
    row['在籍'] = STATUS_ACTIVE;
    state.meibo.push(row);
    await persist();
    closeModal();
    showToast(`${name} さんを追加しました（ID: ${row['生徒ID']}）。`);
    renderAll();
  });
}

/* ========== 一覧TSVコピー ========== */
function openTsvCopyModal() {
  const cols = visibleCols();
  openModal('一覧をコピー（他ツール・Excelへ貼り付け用）', `
    <div class="hint">現在の絞り込み・並び順のまま、選んだ列をタブ区切りでコピーします。座席表や懇談ツールへの貼り付けに使えます。</div>
    <div style="display:flex;gap:12px;flex-wrap:wrap;margin:12px 0">
      ${cols.map(c => `<label><input type="checkbox" data-copycol="${c.key}" ${['_class', '_num', '生徒氏名'].includes(c.key) ? 'checked' : ''}> ${escapeHtml(c.label)}</label>`).join('')}
    </div>
    <label style="font-size:0.9em"><input type="checkbox" id="tcHeader" checked> 見出し行を含める</label>
    <div class="modal-actions"><button class="btn btn-primary" id="tcCopy">コピーする</button></div>`);
  $('tcCopy').addEventListener('click', async () => {
    const selected = [...els.modalBody.querySelectorAll('[data-copycol]:checked')].map(cb => cb.dataset.copycol);
    if (selected.length === 0) { showToast('列を選んでください。', true); return; }
    const colDefs = cols.filter(c => selected.includes(c.key));
    const rows = sortStudents(filteredStudents());
    const lines = [];
    if ($('tcHeader').checked) lines.push(colDefs.map(c => c.label).join('\t'));
    for (const s of rows) lines.push(colDefs.map(c => cellValue(s, c.key)).join('\t'));
    try {
      await navigator.clipboard.writeText(lines.join('\n'));
      closeModal();
      showToast(`${rows.length}行をコピーしました。`);
    } catch (e) {
      showToast('コピーに失敗しました。ブラウザの許可を確認してください。', true);
    }
  });
}

/* ========== 集計 ========== */
function openSummaryModal() {
  const active = activeStudents();
  const inGrade = active.filter(s => { const g = gradeOf(s); return g >= 1 && g <= 3; });
  const countBy = (list, fn) => {
    const m = new Map();
    for (const s of list) { const k = fn(s) || '（未入力）'; m.set(k, (m.get(k) || 0) + 1); }
    return [...m.entries()].sort((a, b) => String(a[0]).localeCompare(String(b[0]), 'ja', { numeric: true }));
  };
  const table = entries => `<table class="summary-table">${entries.map(([k, v]) => `<tr><td>${escapeHtml(String(k))}</td><td>${v}名</td></tr>`).join('')}</table>`;
  const byClass = countBy(inGrade, s => `${gradeOf(s)}年${currentClass(s) || '?'}組`);
  const byClub = countBy(inGrade, s => s['クラブ']);
  const byCommittee = countBy(inGrade.filter(s => s['委員会']), s => s['委員会']);
  const byMonth = countBy(inGrade.filter(s => /^\d{4}-\d{2}/.test(s['生年月日'] || '')), s => parseInt(s['生年月日'].slice(5, 7), 10) + '月');
  const support = inGrade.filter(s => s['支援計画'] === 'あり');
  openModal(`集計（R${state.nendo}年度・在籍のみ）`, `
    <div class="summary-block"><h3>学年・クラス別人数</h3>${table(byClass)}</div>
    <div class="summary-block"><h3>クラブ別</h3>${table(byClub)}</div>
    <div class="summary-block"><h3>委員会別</h3>${table(byCommittee)}</div>
    <div class="summary-block"><h3>月別誕生日</h3>${table(byMonth)}</div>
    <div class="summary-block"><h3>支援計画</h3><p>「あり」${support.length}名${els.toggleSensitive.checked && support.length > 0 ? '：' + support.map(s => escapeHtml(s['生徒氏名'])).join('、') : '（氏名の表示は「個人情報・支援の列を表示」をONに）'}</p></div>`);
}

/* ========== 卒業処理 ========== */
function openGraduateModal() {
  const targets = activeStudents().filter(s => gradeOf(s) === 3);
  if (targets.length === 0) { showToast('3年生の在籍生徒がいません。', true); return; }
  openConfirm('卒業処理', `R${state.nendo}年度の3年生 ${targets.length}名の在籍を「卒業」に変更します。よろしいですか？`, async () => {
    for (const s of targets) s['在籍'] = '卒業';
    await persist();
    showToast(`${targets.length}名を卒業に変更しました。`);
    renderAll();
  });
}

/* ========== カルテ印刷 ========== */
function openPrintModal() {
  openModal('カルテ印刷', `
    <p>この生徒の基本情報・記録・成績を1枚に印刷します。</p>
    <label><input type="checkbox" id="prSens"> 🔒 個人情報（住所・連絡先・健康・支援）を含める</label>
    <div class="modal-actions"><button class="btn btn-primary" id="prGo">印刷する</button></div>`);
  $('prGo').addEventListener('click', () => {
    const withSens = $('prSens').checked;
    closeModal();
    printKarte(withSens);
  });
}

function printKarte(withSens) {
  const s = findStudent(ui.karteId);
  if (!s) return;
  const id = String(s['生徒ID']);
  const skipCols = withSens ? [] : SENSITIVE_COLS;
  let html = '<div class="field-grid">';
  for (const h of state.meiboHeaders) {
    if (skipCols.includes(h)) continue;
    html += `<div class="field"><label>${escapeHtml(h)}</label><div>${escapeHtml(String(s[h] ?? '')) || '—'}</div></div>`;
  }
  html += '</div>';
  const recs = state.records.filter(r => String(r['生徒ID']) === id).sort((a, b) => String(b['日付']).localeCompare(String(a['日付'])));
  const visibleRecs = withSens ? recs : recs.filter(r => !['指導', '支援', '健康'].includes(String(r['種別'])));
  html += `<h3 style="margin-top:16px;color:var(--primary)">記録${withSens ? '' : '（指導・支援・健康の記録は含めていません）'}</h3>`;
  html += visibleRecs.length === 0 ? '<p>—</p>' : visibleRecs.map(r =>
    `<div class="record-item"><div class="record-head"><span>${escapeHtml(r['日付'])}</span><span class="record-type">${escapeHtml(r['種別'])}</span></div><p>${escapeHtml(r['内容'])}</p></div>`).join('');
  const grs = state.grades.filter(r => String(r['生徒ID']) === id);
  html += '<h3 style="margin-top:16px;color:var(--primary)">成績</h3>';
  if (grs.length === 0) html += '<p>—</p>';
  else {
    html += '<table class="summary-table"><tr><th>年度</th><th>学期</th><th>テスト</th><th>教科</th><th>点数</th></tr>' +
      grs.map(r => `<tr><td>R${escapeHtml(r['年度'])}</td><td>${escapeHtml(r['学期'])}</td><td>${escapeHtml(r['テスト名'])}</td><td>${escapeHtml(r['教科'])}</td><td>${escapeHtml(r['点数'])}</td></tr>`).join('') + '</table>';
  }
  const savedTab = ui.karteTab;
  els.karteBody.innerHTML = html;
  document.body.classList.add('printing-karte');
  const restore = () => {
    document.body.classList.remove('printing-karte');
    ui.karteTab = savedTab;
    renderKarte();
    window.removeEventListener('afterprint', restore);
  };
  window.addEventListener('afterprint', restore);
  window.print();
}

/* ========== トースト ========== */
let toastTimer = null;
function showToast(msg, isError) {
  els.toast.textContent = msg;
  els.toast.className = 'toast show' + (isError ? ' toast-error' : '');
  els.toast.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { els.toast.className = 'toast'; }, 3000);
}

init();
