/* app.js — v3.16 (Upstream Industry Avg Ranking)
 * 修正內容：
 *  A) 右側「下游產業」正確讀取 DownLinks sheet
 *  B) renderTreemap 支援 Links / DownLinks 兩種資料格式
 *  C) DownLinks 讀取加入安全檢查
 *  D) 維持原本 .US 不顯示、MoM/YoY 美股顯示為 —
 *  E) 群組標題只顯示平均值，不顯示幾檔
 *  F) 若個股格子太小，小到放不下文字，則直接不呈現該檔個股
 *  G) 節點 tooltip / 點擊查詢
 *  H) 上游 Treemap 改為依 Revenue 的「產業別」分群
 *  I) 上游 Treemap 依各類股平均營收表現排序，只保留前 GROUP_KEEP_MAX 個類股
 */

const URL_VER = new URLSearchParams(location.search).get('v') || Date.now();
const XLSX_FILE = new URL(`./data.xlsx?v=${URL_VER}`, location.href).toString();
const REVENUE_SHEET = 'Revenue';
const LINKS_SHEET   = 'Links';
const DOWNLINKS_SHEET = 'DownLinks';
const NEWHIGH_SHEET   = '創新高';

const CODE_FIELDS = ['個股','代號','股票代碼','股票代號','公司代號','證券代號'];
const NAME_FIELDS = ['名稱','公司名稱','證券名稱'];
const COL_MAP = {};
const AMOUNT_COL_MAP = {}; // 各月份營收金額欄位，例如 202505 -> 202505月合併營收(千)



// ===== 可調參數 =====
const HEADER_H = 22;
const GROUP_KEEP_MAX = 7;
const DOWN_GROUP_STOCK_KEEP_MAX = 18;
const GROUP_WEIGHT_MODE = 'RANK';
const RANK_WEIGHT_MIN = 1.3;
const RANK_WEIGHT_MAX = 1.8;

// ===== 上游類股篩選規則 =====
// true  = 上游只顯示平均值 > 0 的類股
// false = 上游只依平均值排序，允許負值類股進榜
const UPSTREAM_ONLY_POSITIVE = false;

const ENABLE_NODE_CLICK = false;    // 點方塊可重新查詢
const MIN_RENDER_W = 75;           // 個股最小寬度（小於則不顯示）
const MIN_RENDER_H = 20;           // 個股最小高度（小於則不顯示）
const MIN_RENDER_AREA = 400;       // 個股最小面積（小於則不顯示）
const NEWHIGH_COLLAPSE_AFTER = 0; // 營收創新高表格，預設先顯示前 15 檔
// ===== 右側概念股熱力圖保底顯示設定 =====
// 避免 100% 畫面寬度下，所有個股方塊都被判定太小而整張圖消失
const DOWN_MIN_RENDER_W_FLOOR = 38;
const DOWN_MIN_RENDER_H_FLOOR = 14;
const DOWN_MIN_RENDER_AREA_FLOOR = 160;
const DOWN_FALLBACK_KEEP_PER_GROUP = 2;


let revenueRows = [], linksRows = [], downRows = [], downRowsRaw = [], months = [];
let newHighSheetRows = [];
let byCode = new Map();
let byName = new Map();
let linksByUp = new Map();
let linksByDown = new Map();
let currentDownstreamEdges = [];
let currentMonth = '';
let currentMetric = '';
let currentColorMode = '';
let currentSelectedConcept = '';



let upstreamHJ = [];   // 左邊相同產業分類專用（DownLinks G/H/I）
let downstreamHJ = []; // 右邊維持原本

function z(s){ return String(s == null ? '' : s); }
function toHalfWidth(str){ return z(str).replace(/[０-９Ａ-Ｚａ-ｚ]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0)); }
function normText(s){ return z(s).replace(/[\u200B-\u200D\uFEFF]/g,'').replace(/[\u3000]/g,' ').replace(/\s+/g,' ').trim(); }
function normCode(s){ return toHalfWidth(z(s)).replace(/[\u200B-\u200D\uFEFF]/g,'').replace(/\s+/g,'').trim(); }
function displayPct(v){ if(v == null || !isFinite(v)) return '—'; const s = v.toFixed(1) + '%'; return v > 0 ? ('+' + s) : s; }
function colorFor(v, mode){ if(v == null || !isFinite(v)) return '#0f172a'; const t = Math.min(1, Math.abs(v)/80); const alpha = 0.25 + 0.35*t; const good = (mode === 'greenPositive'); const pos = good ? '156,163,175' : '59,130,246'; const neg = good ? '59,130,246' : '156,163,175'; const rgb = (v >= 0) ? pos : neg; return `rgba(${rgb},${alpha})`; }
function safe(s){ return z(s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
function isUSCode(code){ return /\.US$/i.test(String(code || '').trim()); }

function triggerUnifiedQuery(){
  const input = document.querySelector('#stockInput');
  if (input) {
    // 讓其他可能監聽 input/change 的腳本也收到更新
    input.dispatchEvent(new Event('input', { bubbles: true }));
    input.dispatchEvent(new Event('change', { bubbles: true }));
  }

  const btn = document.querySelector('#runBtn');
  if (btn) {
    // 關鍵：模擬按下查詢按鈕，讓其他檔案綁在 click 的邏輯一起跑
    btn.click();
    return;
  }

  // 保底
  handleRun();
}

function interceptQueryControlsEnter(){
  const targets = [
    '#stockInput',
    '#monthSelect',
    '#metricSelect',
    '#colorMode'   // 如果頁面上沒有這個元素也沒關係
  ];

  const onEnter = (e) => {
    if (e.key !== 'Enter') return;
    if (e.isComposing || e.keyCode === 229) return;

    e.preventDefault();
    e.stopPropagation();
    if (typeof e.stopImmediatePropagation === 'function') {
      e.stopImmediatePropagation();
    }

    // 讓下拉選單先失焦，避免 Enter 只是在開/關選單
    if (typeof e.currentTarget?.blur === 'function') {
      e.currentTarget.blur();
    }

    // 下一個 frame 再查詢，確保 select 的值已更新
    requestAnimationFrame(() => {
      triggerUnifiedQuery();
    });
  };

  targets.forEach(selector => {
    const el = document.querySelector(selector);
    if (!el) return;

    // 用 capture=true，盡量比其他外部腳本更早攔截
    el.addEventListener('keydown', onEnter, true);
    el.addEventListener('keypress', onEnter, true);
  });
}

window.addEventListener('DOMContentLoaded', async () => {
  try {
    await loadWorkbook();
    initControls();
    renderNewHighSummary();

    // 改成：攔截所有查詢欄位的 Enter
    interceptQueryControlsEnter();

  } catch (e) {
    console.error(e);
    alert('載入失敗：' + e.message);
  }

  document.querySelector('#runBtn')?.addEventListener('click', (e) => {
    e.preventDefault?.();
    handleRun();
  });
});



async function loadWorkbook(){
  const res = await fetch(XLSX_FILE, { cache:'no-store' });
  if (!res.ok) throw new Error('讀取 data.xlsx 失敗 HTTP ' + res.status);

  const buf = await res.arrayBuffer();
  const wb  = XLSX.read(buf, { type:'array' });
  console.log('[工作表名稱]', wb.SheetNames);

  const wsRev = wb.Sheets[REVENUE_SHEET];
  const wsLinks = wb.Sheets[LINKS_SHEET];
  const wsDown = wb.Sheets[DOWNLINKS_SHEET];
  const wsNewHigh = wb.Sheets[NEWHIGH_SHEET];

  if (!wsRev || !wsLinks) throw new Error('找不到必要工作表 Revenue 或 Links');

  const rowsHeaderFirst = XLSX.utils.sheet_to_json(wsRev, { header: 1, blankrows: false });
  const headerRow = Array.isArray(rowsHeaderFirst) && rowsHeaderFirst.length > 0 ? rowsHeaderFirst[0] : [];
  const found = new Set();

  // 先清空，避免舊資料殘留
  for (const k of Object.keys(COL_MAP)) delete COL_MAP[k];

  console.log('[Revenue headerRow]', headerRow);

for (const rawHeader of headerRow) {
  if (!rawHeader) continue;

  const h = normText(String(rawHeader));
  console.log('[檢查欄名]', h);

  // 先抓年月：支援 2025/3、2025年3月、2025-03、202503
  let ymMatch =
    h.match(/(20\d{2})[\/年\-]?\s*(\d{1,2})\s*月?/) ||
    h.match(/(20\d{2})(\d{2})/);

  if (!ymMatch) continue;

  const ym = ymMatch[1] + String(ymMatch[2]).padStart(2, '0');

  // 判斷欄位是 YoY 還是 MoM
  const isYoY = /年增|年成長|YoY/i.test(h);
  const isMoM = /月增|月變動|MoM/i.test(h);

  if (isYoY) {
    (COL_MAP[ym] ??= {}).YoY = rawHeader;
    found.add(ym);
    console.log(`[COL_MAP][${ym}].YoY =`, rawHeader);
    continue;
  }

  if (isMoM) {
    (COL_MAP[ym] ??= {}).MoM = rawHeader;
    found.add(ym);
    console.log(`[COL_MAP][${ym}].MoM =`, rawHeader);
    continue;
  }

  // 新增：抓每個月份的「營收金額」欄位
  // 例如 202505月合併營收(千)、2025年5月營收、2025/05營收
  // 排除 YoY / MoM，避免抓到成長率欄位
  const isRevenueAmount =
    /營收/i.test(h) &&
    !/年增|年成長|YoY|月增|月變動|MoM/i.test(h);

  if (isRevenueAmount) {
    AMOUNT_COL_MAP[ym] = rawHeader;
    found.add(ym);
    console.log(`[AMOUNT_COL_MAP][${ym}] =`, rawHeader);
    continue;
  }
}

  months = Array.from(found).sort((a, b) => b.localeCompare(a));

  console.log('[months]', months);
  console.log('[COL_MAP]', COL_MAP);

  

  revenueRows = XLSX.utils.sheet_to_json(wsRev,   { defval:null });
  linksRows   = XLSX.utils.sheet_to_json(wsLinks, { defval:null });
  downRows    = wsDown ? XLSX.utils.sheet_to_json(wsDown, { defval:null }) : [];
  downRowsRaw = wsDown ? XLSX.utils.sheet_to_json(wsDown, { header:1, defval:'', blankrows:false }) : [];

  newHighSheetRows = wsNewHigh
    ? XLSX.utils.sheet_to_json(wsNewHigh, { header: 1, defval: '', blankrows: false, raw: false })
    : [];

  byCode.clear();
  byName.clear();

  const sample = revenueRows[0] || {};
  const codeKeyName = CODE_FIELDS.find(k => k in sample) || CODE_FIELDS[0];
  const nameKeyName = NAME_FIELDS.find(k => k in sample) || NAME_FIELDS[0];

  for (const r of revenueRows) {
    const code = normCode(String(r[codeKeyName]).replace(/\u3000/g, '').replace(/[\u200B-\u200D\uFEFF]/g, '').trim());
    const name = normText(r[nameKeyName]);
    if (code) byCode.set(code, r);
    if (name) byName.set(name, r);
  }

linksByUp.clear();
linksByDown.clear();

// ===== Links（保留原本資料結構，避免其他功能受影響）=====
for (const e of linksRows) {
  const A = normCode(e['上游代號']);
  const B = normCode(e['下游代號']);
  const C = normText(e['關係類型']);

  if (A && B && C) {
    if (!linksByUp.has(A)) linksByUp.set(A, []);
    linksByUp.get(A).push(e);

    if (!linksByDown.has(B)) linksByDown.set(B, []);
    linksByDown.get(B).push(e);
  }
}

// ===== DownLinks（左邊「相同產業分類」熱力圖專用：直接讀 G/H/I 欄）=====
// G index=6：上游代號_熱力圖上
// H index=7：下游代號_熱力圖上
// I index=8：關係類型_熱力圖上（你現在的分類名稱）
upstreamHJ = [];
for (const row of downRowsRaw.slice(1)) { // 略過標題列
  const up = normCode(row[6]);    // G
  const down = normCode(row[7]);  // H
  const type = normText(row[8]);  // I

  if (up && down && type) {
    upstreamHJ.push({
      '上游代號': up,
      '下游代號': down,
      '關係類型': type
    });
  }
}

// ===== DownLinks（右邊下游）=====
downstreamHJ = [];
for (const row of downRows) {
  const up = normCode(row['上游代號']);
  const down = normCode(row['下游代號']);
  const type = normText(row['關係類型']);

  if (up && down && type) {
    downstreamHJ.push({
      '上游代號': up,
      '下游代號': down,
      '關係類型': type
    });
  }
}

console.log("Links 筆數 =", linksRows.length);
console.log("左邊相同產業分類 DownLinks(GHI) 筆數 =", upstreamHJ.length);
console.log("右邊 DownLinks(ABC) 筆數 =", downstreamHJ.length);
}

function initControls(){
  const sel = document.querySelector('#monthSelect');
  if (!sel) {
    console.warn('[initControls] 找不到 #monthSelect');
    return;
  }

  sel.innerHTML = '';

  if (!months.length) {
    console.error('[initControls] months 是空的，無法建立月份下拉選單');
    const o = document.createElement('option');
    o.value = '';
    o.textContent = '無可用月份';
    sel.appendChild(o);
    return;
  }

  for (const m of months) {
    const o = document.createElement('option');
    o.value = m;
    o.textContent = `${m.slice(0,4)}年${m.slice(4,6)}月`;
    sel.appendChild(o);
  }

  sel.value = months[0];
  console.log('[initControls] monthSelect 預設值 =', sel.value);
}

function getMetricValue(row, month, metric){
  if (!row || !month || !metric) return null;

  const codeOfRow = normCode(
    row['個股'] || row['代號'] || row['股票代碼'] ||
    row['股票代號'] || row['公司代號'] || row['證券代號'] || ''
  );

  if (isUSCode(codeOfRow)) return null;

  const col = (COL_MAP[month] || {})[metric];
  if (!col) return null;

  let v = row[col];
  if (v == null || v === '') return null;

  if (typeof v === 'string') v = v.replace('%','').replace('％','').trim();
  v = Number(v);

  return Number.isFinite(v) ? v : null;
}

function shiftMonth(ym, diff){
  if (!ym || ym.length !== 6) return '';

  const y = Number(ym.slice(0, 4));
  const m = Number(ym.slice(4, 6));

  if (!Number.isFinite(y) || !Number.isFinite(m)) return '';

  const total = y * 12 + (m - 1) + diff;
  const ny = Math.floor(total / 12);
  const nm = total % 12 + 1;

  return String(ny) + String(nm).padStart(2, '0');
}

function getRevenueAmountValue(row, month){
  if (!row || !month) return null;

  const codeOfRow = normCode(
    row['個股'] || row['代號'] || row['股票代碼'] ||
    row['股票代號'] || row['公司代號'] || row['證券代號'] || ''
  );

  // 美股不納入台股營收合計
  if (isUSCode(codeOfRow)) return null;

  const col = AMOUNT_COL_MAP[month];
  if (!col) return null;

  let v = row[col];

  if (v == null || v === '') return null;

  if (typeof v === 'string') {
    v = v.replace(/[%％,\s]/g, '').trim();
  }

  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

function getGroupAggregateRevenuePerformance(rows, month, metric){
  if (!Array.isArray(rows) || !rows.length || !month || !metric) return null;

  const baseMonth = metric === 'YoY'
    ? shiftMonth(month, -12)
    : shiftMonth(month, -1);

  if (!baseMonth) return null;

  let currentSum = 0;
  let baseSum = 0;
  let validCount = 0;

  for (const row of rows) {
    const currentRevenue = getRevenueAmountValue(row, month);
    const baseRevenue = getRevenueAmountValue(row, baseMonth);

    if (!Number.isFinite(currentRevenue)) continue;
    if (!Number.isFinite(baseRevenue)) continue;
    if (baseRevenue === 0) continue;

    currentSum += currentRevenue;
    baseSum += baseRevenue;
    validCount += 1;
  }

  if (!validCount || baseSum === 0) return null;

  return (currentSum / baseSum - 1) * 100;
}



function shouldSkipTreemapValue(v){
  // 無效值、空值、找不到資料，直接不放進熱力圖
  if (!Number.isFinite(v)) return true;

  // 因為畫面只顯示到小數第 1 位
  // 所以 -0.04、0.03 這種最後會顯示成 0.0% 的值，也直接省略
  return Math.abs(v) < 0.05;
}



// ===== 新增：決定 Treemap 分群名稱 =====
// 上游：改用 Revenue 的「產業別」分群
// 下游：維持原本用 Links / DownLinks 的「關係類型」分群
function getTreemapGroupName(svgId, edge, row){
  if (svgId === 'upTreemap') {
    // 左邊改成依 DownLinks 的 I 欄分群
    return normText(edge['關係類型'] || edge['type'] || '相同產業分類');
  }
  return normText(edge['關係類型'] || edge['type'] || '未分類');
}

// ===== 新增：決定 Treemap 要保留哪些群組 =====
// 下游：維持原本依群組股票數量排序
// 上游：改為依「類股平均值」由高到低排序，最多保留 GROUP_KEEP_MAX 個類股
function selectTreemapGroups(svgId, summaries){
  // 右邊概念股：不要只挑檔數最多的群組
  // 改成優先挑「平均表現較佳」且「檔數不要過度膨脹」的群組
  if (svgId !== 'upTreemap') {
    return [...summaries]
      .filter(g => Number.isFinite(g.avg))
      .sort((a, b) => {
        // 先比平均表現
        if (b.avg !== a.avg) return b.avg - a.avg;

        // 平均表現相同時，再比檔數
        return b.list.length - a.list.length;
      })
      .slice(0, GROUP_KEEP_MAX);
  }

  let arr = summaries.filter(g => Number.isFinite(g.avg));

  if (UPSTREAM_ONLY_POSITIVE) {
    arr = arr.filter(g => g.avg > 0);
  }

  return arr
    .sort((a, b) => b.avg - a.avg)
    .slice(0, GROUP_KEEP_MAX);
}


function treemapSortByPerformance(a, b){
  const aHasChildren = !!a.children;
  const bHasChildren = !!b.children;

  // 群組層：依群組平均值由高到低排序
  if (aHasChildren || bHasChildren) {
    const av = Number.isFinite(a.data?.avg) ? a.data.avg : -Infinity;
    const bv = Number.isFinite(b.data?.avg) ? b.data.avg : -Infinity;

    if (bv !== av) return bv - av;

    return String(a.data?.name || '').localeCompare(
      String(b.data?.name || ''),
      'zh-Hant',
      { numeric: true }
    );
  }

  // 個股層：依目前查詢月份的 YoY / MoM 由高到低排序
  const av = Number.isFinite(a.data?.raw) ? a.data.raw : -Infinity;
  const bv = Number.isFinite(b.data?.raw) ? b.data.raw : -Infinity;

  if (bv !== av) return bv - av;

  return String(a.data?.code || '').localeCompare(
    String(b.data?.code || ''),
    'zh-Hant',
    { numeric: true }
  );
}


function getTreemapLeafBase(raw, svgId){
  const minBase = (svgId === 'upTreemap') ? 1.2 : 0.9;

  if (!Number.isFinite(raw)) return minBase;

  // 正成長：數值越大，格子越大
  if (raw > 0) {
    const cap = (svgId === 'downTreemap') ? 180 : 220;
    const exponent = (svgId === 'downTreemap') ? 0.42 : 0.50;
    const capped = Math.min(raw, cap);

    return minBase + Math.pow(capped + 1, exponent);
  }

  // 0 附近：給基本大小
  if (raw === 0) {
    return minBase;
  }

  // 負成長：越負，格子越小
  // 例如 -5% 仍有一定大小，-50%、-100% 會明顯縮小
  const negAbs = Math.abs(raw);
  const shrink = 1 / Math.pow(1 + negAbs / 20, 0.75);

  return Math.max(0.18, minBase * shrink);
}




function handleRun(){
  const raw = document.querySelector('#stockInput').value;
  const month = (document.querySelector('#monthSelect')?.value) || '';
  const metric = (document.querySelector('#metricSelect')?.value) || 'MoM';
  const colorMode = (document.querySelector('#colorMode')?.value) || 'redPositive';
  console.log('[handleRun] raw =', raw);
  console.log('[handleRun] month =', month);
  console.log('[handleRun] metric =', metric);
  console.log('[handleRun] colorMode =', colorMode);

  if (!month) {
    alert('月份尚未載入，請先檢查 Revenue 工作表欄名格式是否有抓到月份');
    return;
  }

  if (!raw || !raw.trim()) {
    alert('請輸入股票代號或公司名稱');
    return;
  }

  let codeKey = normCode(raw);
  let rowSelf = byCode.get(codeKey);

  if (!rowSelf) {
    const nameQ = normText(raw);
    rowSelf = byName.get(nameQ) || revenueRows.find(r => normText(r['名稱'] || r['公司名稱'] || r['證券名稱'] || '').startsWith(nameQ));
    if (rowSelf) {
      codeKey = normCode(
        rowSelf['個股'] ?? rowSelf['代號'] ?? rowSelf['股票代碼'] ??
        rowSelf['股票代號'] ?? rowSelf['公司代號'] ?? rowSelf['證券代號']
      );
    }
  }

  if (!rowSelf) {
    alert('找不到此代號/名稱');
    return;
  }

  try {
    const codeLabel = (rowSelf['個股'] || rowSelf['代號'] || rowSelf['股票代碼'] || rowSelf['股票代號'] || rowSelf['公司代號'] || rowSelf['證券代號'] || '').trim();
    const nameLabel = (rowSelf['名稱'] || rowSelf['公司名稱'] || rowSelf['證券名稱'] || '').trim();
    const extra = `${month.slice(0,4)}/${month.slice(4,6)} · ${metric}`;
    if (window.setResultChipLink) window.setResultChipLink(codeLabel, nameLabel, extra);
  } catch (_) {}

let upstreamEdges = upstreamHJ.filter(e => e['下游代號'] === codeKey);
let downstreamEdges = downstreamHJ.filter(e => e['上游代號'] === codeKey);

// 左邊相同產業股加入查詢個股本身
const upstreamGroupNames = [...new Set(
  upstreamEdges
    .map(e => normText(e['關係類型'] || e['type'] || ''))
    .filter(Boolean)
)];

const selfAlreadyInUpstream = upstreamEdges.some(e =>
  normCode(e['上游代號']) === codeKey
);

if (!selfAlreadyInUpstream) {
  const fallbackGroup = normText(rowSelf['產業別'] || '相同產業分類');
  const groupsToAdd = upstreamGroupNames.length ? upstreamGroupNames : [fallbackGroup];

  for (const groupName of groupsToAdd) {
    upstreamEdges.push({
      '上游代號': codeKey,
      '下游代號': codeKey,
      '關係類型': groupName,
      '__self': true
    });
  }
}

// 右邊概念股加入查詢個股本身
const downstreamGroupNames = [...new Set(
  downstreamEdges
    .map(e => normText(e['關係類型'] || e['type'] || ''))
    .filter(Boolean)
)];

const selfAlreadyInDownstreamGroups = new Set(
  downstreamEdges
    .filter(e => normCode(e['下游代號']) === codeKey)
    .map(e => normText(e['關係類型'] || e['type'] || ''))
);

for (const groupName of downstreamGroupNames) {
  if (selfAlreadyInDownstreamGroups.has(groupName)) continue;

  downstreamEdges.push({
    '上游代號': codeKey,
    '下游代號': codeKey,
    '關係類型': groupName,
    '__self': true
  });
}


  downstreamEdges = downstreamEdges.filter(e => !String(e['下游代號']).endsWith('.US'));

  console.log('查詢代號 =', codeKey);
  console.log('上游筆數 =', upstreamEdges.length, upstreamEdges);
  console.log('下游筆數 =', downstreamEdges.length, downstreamEdges);
  
currentDownstreamEdges = downstreamEdges;
currentMonth = month;
currentMetric = metric;
currentColorMode = colorMode;
currentSelectedConcept = '';

renderConceptNote(rowSelf, downstreamEdges);
renderConceptStockList('', downstreamEdges, month, metric);


requestAnimationFrame(() => {
    try {
      renderResultChip(rowSelf, month, metric, colorMode);
      renderTreemap('upTreemap', 'upHint', upstreamEdges, '上游代號', month, metric, colorMode);
    } catch (err) {
      console.error('[handleRun][upTreemap]', err);
    }
  });

  requestAnimationFrame(() => {
    try {
      renderTreemap('downTreemap', 'downHint', downstreamEdges, '下游代號', month, metric, colorMode);
    } catch (err) {
      console.error('[handleRun][downTreemap]', err);
    }
  });
}


function renderResultChip(selfRow, month, metric, colorMode){
  const host = document.querySelector('#resultChip');
  if (!host) {
    console.warn('[renderResultChip] 找不到 #resultChip，略過 result card 繪製');
    return;
  }

  const v = getMetricValue(selfRow, month, metric);
  const bg = colorFor(v, colorMode);

  const showCode = selfRow['個股'] || selfRow['代號'] || selfRow['股票代碼'] || selfRow['股票代號'] || selfRow['公司代號'] || selfRow['證券代號'] || '';
  const showName = selfRow['名稱'] || selfRow['公司名稱'] || selfRow['證券名稱'] || '';

  host.innerHTML = `
    <div class="result-card" style="background:${bg}">
      <div class="row1"><strong>${safe(showCode)}｜${safe(showName)}</strong><span>${month.slice(0,4)}/${month.slice(4,6)} / ${metric}</span></div>
      <div class="row2"><span>${safe(selfRow['產業別'] || '')}</span><span>${displayPct(v)}</span></div>
    </div>`;
}







function renderConceptNote(selfRow, downstreamEdges){
  //ensureConceptNoteStyles();
  const host = document.querySelector('#conceptNote');

  if (!host) {
    console.warn('[renderConceptNote] 找不到 #conceptNote');
    return;
  }

  const showCode =
    selfRow['個股'] ||
    selfRow['代號'] ||
    selfRow['股票代碼'] ||
    selfRow['股票代號'] ||
    selfRow['公司代號'] ||
    selfRow['證券代號'] ||
    '';

  const showName =
    selfRow['名稱'] ||
    selfRow['公司名稱'] ||
    selfRow['證券名稱'] ||
    '';

  const concepts = [...new Set(
    (downstreamEdges || [])
      .map(e => normText(e['關係類型'] || e['type'] || ''))
      .filter(Boolean)
  )];

  if (!concepts.length) {
    host.innerHTML = `
      <div class="concept-note-inner">
        <span class="concept-note-title">概念股補充</span>
        <span class="concept-note-text">
          目前查無 ${safe(showCode)} ${safe(showName)} 的概念股分類資料。
        </span>
      </div>
    `;
    return;
  }

const conceptButtonsHtml = concepts.map(c =>
  `<span role="button" tabindex="0" class="concept-chip-btn" data-concept="${safe(c)}" title="查看 ${safe(c)} 完整概念股名單">${safe(c)}</span>`
).join('、');

host.innerHTML = `
  <div class="concept-note-inner">
    <span class="concept-note-title">概念股補充</span>
    <span class="concept-note-text">概念股清單較多時，熱力圖僅呈現部分代表分類與個股。此個股為 <span class="concept-list">${conceptButtonsHtml}</span> 概念股。</span>
  </div>
`;

  host.querySelectorAll('.concept-chip-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const concept = normText(btn.dataset.concept || '');
      currentSelectedConcept = concept;

      host.querySelectorAll('.concept-chip-btn').forEach(b => {
        b.classList.toggle('active', normText(b.dataset.concept || '') === concept);
      });

      renderConceptStockList(
        concept,
        currentDownstreamEdges,
        currentMonth,
        currentMetric
      );
    });
  });
}


function getRevenueRowCode(row){
  if (!row) return '';

  return normCode(
    row['個股'] ||
    row['代號'] ||
    row['股票代碼'] ||
    row['股票代號'] ||
    row['公司代號'] ||
    row['證券代號'] ||
    ''
  );
}

function getRevenueRowName(row){
  if (!row) return '';

  return normText(
    row['名稱'] ||
    row['公司名稱'] ||
    row['證券名稱'] ||
    ''
  );
}

function buildConceptStockRecords(conceptName, downstreamEdges, month, metric){
  const concept = normText(conceptName);
  if (!concept) return [];

  const seen = new Set();
  const records = [];

  for (const e of downstreamEdges || []) {
    const rel = normText(e['關係類型'] || e['type'] || '');
    if (rel !== concept) continue;

    const code = normCode(e['下游代號'] || e['down'] || '');
    if (!code) continue;
    if (isUSCode(code)) continue;
    if (seen.has(code)) continue;

    seen.add(code);

    const row = byCode.get(code);
    const name = row ? getRevenueRowName(row) : '';
    const industry = row ? normText(row['產業別'] || '') : '';

    const mom = row ? getMetricValue(row, month, 'MoM') : null;
    const yoy = row ? getMetricValue(row, month, 'YoY') : null;

    const sortValue = metric === 'YoY' ? yoy : mom;

    records.push({
      code,
      name,
      industry,
      mom,
      yoy,
      value: sortValue
    });
  }

  records.sort((a, b) => {
    const av = Number.isFinite(a.value) ? a.value : -Infinity;
    const bv = Number.isFinite(b.value) ? b.value : -Infinity;

    if (bv !== av) return bv - av;

    return String(a.code).localeCompare(
      String(b.code),
      'zh-Hant',
      { numeric: true }
    );
  });

  return records;
}


function ensureConceptStockTableStyles(){
  // 每次重新注入，避免舊版 inline style 殘留
  const oldStyle = document.getElementById('conceptStockTableInlineStyle');
  if (oldStyle) oldStyle.remove();

  const style = document.createElement('style');
  style.id = 'conceptStockTableInlineStyle';
  style.textContent = `
    #conceptStockListWrap {
      max-width: 100% !important;
      overflow-x: hidden !important;
    }

    #conceptStockListWrap .concept-stock-table-wrap {
      width: 100% !important;
      max-width: 100% !important;
      overflow-x: hidden !important;
      box-sizing: border-box !important;
    }

    #conceptStockListWrap .concept-stock-table {
      width: 100% !important;
      max-width: 100% !important;
      min-width: 0 !important;
      border-collapse: collapse !important;
      table-layout: fixed !important;
      box-sizing: border-box !important;
    }

    /* 欄寬重新分配：讓整張表剛好塞進右側卡片 */
    #conceptStockListWrap .concept-stock-table col.col-code {
      width: 12% !important;
    }

    #conceptStockListWrap .concept-stock-table col.col-name {
      width: 23% !important;
    }

    #conceptStockListWrap .concept-stock-table col.col-industry {
      width: 27% !important;
    }

    #conceptStockListWrap .concept-stock-table col.col-mom {
      width: 19% !important;
    }

    #conceptStockListWrap .concept-stock-table col.col-yoy {
      width: 19% !important;
    }

    #conceptStockListWrap .concept-stock-table thead th,
    #conceptStockListWrap .concept-stock-table tbody td {
      padding: 9px 6px !important;
      vertical-align: middle !important;
      box-sizing: border-box !important;
      color: #f8fafc !important;
      font-size: 15px !important;
      overflow: hidden !important;
      text-overflow: ellipsis !important;
    }

    #conceptStockListWrap .concept-stock-table thead th {
      color: #dbeafe !important;
      font-weight: 600 !important;
      text-align: left !important;
      white-space: nowrap !important;
      border-bottom: 1px solid rgba(255,255,255,0.10) !important;
    }

    #conceptStockListWrap .concept-stock-table tbody td {
      border-bottom: 1px dashed rgba(255,255,255,0.05) !important;
    }

    #conceptStockListWrap .concept-stock-table thead th + th,
    #conceptStockListWrap .concept-stock-table tbody td + td {
      border-left: 1px solid rgba(255,255,255,0.07) !important;
    }

    #conceptStockListWrap .concept-stock-table tbody td.code,
    #conceptStockListWrap .concept-stock-table tbody td.name,
    #conceptStockListWrap .concept-stock-table tbody td.industry {
      text-align: left !important;
      white-space: nowrap !important;
    }

    #conceptStockListWrap .concept-stock-table thead th.th-num,
    #conceptStockListWrap .concept-stock-table tbody td.num {
      text-align: right !important;
      font-variant-numeric: tabular-nums !important;
      white-space: nowrap !important;
    }

    /* 讓最後一欄 YoY 不要貼到最右邊被切到 */
    #conceptStockListWrap .concept-stock-table thead th:last-child,
    #conceptStockListWrap .concept-stock-table tbody td:last-child {
      padding-right: 10px !important;
    }

    #conceptStockListWrap .stock-link,
    #conceptStockListWrap .stock-link:link,
    #conceptStockListWrap .stock-link:visited,
    #conceptStockListWrap .stock-link:hover,
    #conceptStockListWrap .stock-link:active,
    #conceptStockListWrap .stock-link:focus {
      color: #f8fafc !important;
      text-decoration: none !important;
      background: transparent !important;
      outline: none !important;
      box-shadow: none !important;
    }

    #conceptStockListWrap .stock-link {
      display: inline-block !important;
      cursor: pointer !important;
      line-height: 1.2 !important;
      font-weight: 600 !important;
      max-width: 100% !important;
      overflow: hidden !important;
      text-overflow: ellipsis !important;
      white-space: nowrap !important;
    }
  `;

  document.head.appendChild(style);
}




function renderConceptStockList(conceptName, downstreamEdges, month, metric){
  const host = document.getElementById('conceptStockListWrap');
  const titleEl = document.getElementById('conceptStockListTitle');
  const metaEl = document.getElementById('conceptStockListMeta');

  if (!host) {
    console.warn('[renderConceptStockList] 找不到 #conceptStockListWrap');
    return;
  }
  ensureConceptStockTableStyles();

  if (!conceptName) {
    if (titleEl) titleEl.textContent = '概念股完整名單';
    if (metaEl) metaEl.textContent = '請點選上方「概念股補充」中的概念股名稱';
    host.innerHTML = `
      <div class="concept-stock-empty">
        點選上方概念股名稱後，這裡會顯示該概念股的完整個股名單。
      </div>
    `;
    return;
  }

  const records = buildConceptStockRecords(conceptName, downstreamEdges, month, metric);

  if (titleEl) {
    titleEl.textContent = `${conceptName} 概念股完整名單`;
  }

if (metaEl) {
  const monthLabel = formatMonthLabelForConceptTable(month);

  metaEl.textContent = records.length
    ? `資料月份：${monthLabel}｜共 ${records.length} 檔｜點選代號或名稱可開啟富邦投信 K 線圖`
    : '查無符合資料';
}

  if (!records.length) {
    host.innerHTML = `
      <div class="concept-stock-empty">
        目前查無「${safe(conceptName)}」概念股完整名單。
      </div>
    `;
    return;
  }

  const rowsHtml = records.map(r => {
    const url = getStockPageUrl(r.code);

    return `
      <tr>
        <td class="code">
          <a
            class="stock-link"
            data-code="${safe(r.code)}"
            href="${safe(url)}"
            target="_blank"
            rel="noopener noreferrer"
          >
            ${safe(r.code)}
          </a>
        </td>
        <td class="name">
          <a
            class="stock-link"
            data-code="${safe(r.code)}"
            href="${safe(url)}"
            target="_blank"
            rel="noopener noreferrer"
          >
            ${safe(r.name || '-')}
          </a>
        </td>
        <td class="industry">${safe(r.industry || '-')}</td>
        <td class="num">${displayPct(r.mom)}</td>
        <td class="num">${displayPct(r.yoy)}</td>
      </tr>
    `;
  }).join('');

host.innerHTML = `
  <div class="concept-stock-table-wrap">
    <table class="concept-stock-table">
      <colgroup>
        <col class="col-code">
        <col class="col-name">
        <col class="col-industry">
        <col class="col-mom">
        <col class="col-yoy">
      </colgroup>
      <thead>
          <tr>
            <th class="th-code">代號</th>
            <th class="th-name">名稱</th>
            <th class="th-industry">產業</th>
            <th class="th-num">MoM</th>
            <th class="th-num">YoY</th>
          </tr>
        </thead>
        <tbody>
          ${rowsHtml}
        </tbody>
      </table>
    </div>
  `;
}



// ========= 個股標籤適配 =========
const LabelFit = {
  paddingBase: 8,
  maxFont: 36,
  minFontSoft: 9,
  minFontHard: 8,
  lineHeight: 1.15,

  dynPadding(w, h){
    const m = Math.min(w, h);
    return Math.max(2, Math.min(this.paddingBase, Math.floor(m * 0.08)));
  },

  centerText(el, w, h, p){
    el.setAttribute('text-anchor', 'middle');
    el.setAttribute('dominant-baseline', 'middle');
    el.setAttribute('x', p + Math.max(0, (w - p * 2) / 2));
    el.setAttribute('y', p + Math.max(0, (h - p * 2) / 2));
  },

  ensureClip(gEl, w, h){
    const inset = 2;
    const svg = gEl.ownerSVGElement;
    let defs = svg.querySelector('defs');

    if (!defs) defs = svg.insertBefore(document.createElementNS('http://www.w3.org/2000/svg', 'defs'), svg.firstChild);

    const id = gEl.dataset.clipId || ('clip-' + Math.random().toString(36).slice(2));
    gEl.dataset.clipId = id;

    let clip = svg.querySelector('#' + id);
    if (!clip) {
      clip = document.createElementNS('http://www.w3.org/2000/svg', 'clipPath');
      clip.setAttribute('id', id);
      const r = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
      clip.appendChild(r);
      defs.appendChild(clip);
    }

    const rect = clip.firstChild;
    rect.setAttribute('x', inset);
    rect.setAttribute('y', inset);
    rect.setAttribute('width', Math.max(0, w - inset * 2));
    rect.setAttribute('height', Math.max(0, h - inset * 2));

    gEl.querySelectorAll('text').forEach(t => t.setAttribute('clip-path', `url(#${id})`));
  },

ellipsizeNameToWidth(textEl, maxW){
    const tspans = textEl.querySelectorAll('tspan');
    if (!tspans || tspans.length === 0) return;

    // 只有三行模式（代號 / 名稱 / 漲跌幅）時，第 2 行才是名稱
    if (tspans.length >= 3) {
      const nameTspan = tspans[1];
      let name = nameTspan.textContent || '';

      while (nameTspan.getComputedTextLength() > maxW && name.length > 0) {
        name = name.slice(0, -1);
        nameTspan.textContent = name ? (name + '…') : '';
      }
      return;
    }

    // 保底：舊版單行「代號 名稱」格式
    const t1 = tspans[0];
    const full = t1.textContent || '';
    const m = full.match(/^(\d{4})\s*(.*)$/);
    let code = '', name = full;

    if (m) {
      code = m[1];
      name = m[2] || '';
    }

    t1.textContent = code + (name ? (' ' + name) : '');

    while (t1.getComputedTextLength() > maxW && name.length > 0) {
      name = name.slice(0, -1);
      t1.textContent = code + (name ? (' ' + name + '…') : '');
    }
  },

  canFit(textEl, w, h){
    const p = this.dynPadding(w, h);
    const targetW = Math.max(1, w - p * 2), targetH = Math.max(1, h - p * 2);
    const code = textEl.dataset.code || '';
    const name = textEl.dataset.name || '';
    const pct = textEl.dataset.pct || '';

  const layouts = [
    () => [code, name, pct].filter(Boolean),
    () => [code, pct].filter(Boolean)
  ];


    const k = 0.12;
    const areaFont = Math.sqrt(targetW * targetH) * k;
    const logicalMax = Math.min(this.maxFont, Math.floor(targetH * 0.5));

    for (const L of layouts) {
      while (textEl.firstChild) textEl.removeChild(textEl.firstChild);

      L().forEach(s => {
        const t = document.createElementNS('http://www.w3.org/2000/svg', 'tspan');
        t.textContent = s;
        textEl.appendChild(t);
      });

      let f = Math.max(this.minFontHard, Math.min(logicalMax, Math.floor(areaFont)));
      textEl.setAttribute('font-size', f);
      this.centerText(textEl, w, h, p);
      this.ellipsizeNameToWidth(textEl, targetW);

      let guard = 0;
      while (guard++ < 60) {
        const bb = textEl.getBBox();
        const sW = targetW / Math.max(1, bb.width), sH = targetH / Math.max(1, bb.height);
        const s = Math.min(sW, sH, 1);
        const next = Math.max(this.minFontHard, Math.floor(f * s));
        if (next < f) {
          f = next;
          textEl.setAttribute('font-size', f);
          this.centerText(textEl, w, h, p);
          continue;
        }
        break;
      }

      const tsp = textEl.querySelectorAll('tspan');
      const n = Math.max(1, tsp.length);
      const offsetEm = -((n - 1) * this.lineHeight / 2);
      tsp.forEach((t, i) => {
        t.setAttribute('x', textEl.getAttribute('x'));
        t.setAttribute('dy', i === 0 ? `${offsetEm}em` : `${this.lineHeight}em`);
      });

      const box = textEl.getBBox();
      if (box.width <= targetW + 0.1 && box.height <= targetH + 0.1) {
        return true;
      }
    }

    return false;
  },

  fitBlock(textEl, w, h){
    const p = this.dynPadding(w, h);
    const targetW = Math.max(1, w - p * 2), targetH = Math.max(1, h - p * 2);
    const code = textEl.dataset.code || '';
    const name = textEl.dataset.name || '';
    const pct = textEl.dataset.pct || '';

    const layouts = [
      () => [code, name, pct].filter(Boolean),
      () => [code, pct].filter(Boolean)
    ];


    const k = 0.12;
    const areaFont = Math.sqrt(targetW * targetH) * k;
    const logicalMax = Math.min(this.maxFont, Math.floor(targetH * 0.5));

    for (const L of layouts) {
      while (textEl.firstChild) textEl.removeChild(textEl.firstChild);

      L().forEach(s => {
        const t = document.createElementNS('http://www.w3.org/2000/svg', 'tspan');
        t.textContent = s;
        textEl.appendChild(t);
      });

      let f = Math.max(this.minFontHard, Math.min(logicalMax, Math.floor(areaFont)));
      textEl.setAttribute('font-size', f);
      this.centerText(textEl, w, h, p);
      this.ellipsizeNameToWidth(textEl, targetW);

      let guard = 0;
      while (guard++ < 60) {
        const bb = textEl.getBBox();
        const sW = targetW / Math.max(1, bb.width), sH = targetH / Math.max(1, bb.height);
        const s = Math.min(sW, sH, 1);
        const next = Math.max(this.minFontHard, Math.floor(f * s));
        if (next < f) {
          f = next;
          textEl.setAttribute('font-size', f);
          this.centerText(textEl, w, h, p);
          continue;
        }
        break;
      }

      const tsp = textEl.querySelectorAll('tspan');
      const n = Math.max(1, tsp.length);
      const offsetEm = -((n - 1) * this.lineHeight / 2);
      tsp.forEach((t, i) => {
        t.setAttribute('x', textEl.getAttribute('x'));
        t.setAttribute('dy', i === 0 ? `${offsetEm}em` : `${this.lineHeight}em`);
      });

      const box = textEl.getBBox();
      if (box.width <= targetW + 0.1 && box.height <= targetH + 0.1) {
        textEl.removeAttribute('display');
        return true;
      }
    }

    textEl.setAttribute('display', 'none');
    return false;
  }
};

// ========= 群組標題 =========
const GroupTitleFit = {
  minFont: 5,
  lineHeight: 1.12,
  inset: 4,
  k: 0.12,

  ensureHeaderClip(svg, gEl, d, headerH){
    const id = gEl.dataset.headerClipId || ('hclip-' + Math.random().toString(36).slice(2));
    gEl.dataset.headerClipId = id;
    let defs = svg.querySelector('defs');

    if (!defs) defs = svg.insertBefore(document.createElementNS('http://www.w3.org/2000/svg', 'defs'), svg.firstChild);

    let clip = svg.querySelector('#' + id);
    if (!clip) {
      clip = document.createElementNS('http://www.w3.org/2000/svg', 'clipPath');
      clip.setAttribute('id', id);
      clip.setAttribute('clipPathUnits', 'userSpaceOnUse');
      const r = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
      clip.appendChild(r);
      defs.appendChild(clip);
    }

    const r = clip.firstChild;
    const w = Math.max(0, d.x1 - d.x0), h = Math.max(0, headerH);
    r.setAttribute('x', d.x0 + this.inset);
    r.setAttribute('y', d.y0 + this.inset);
    r.setAttribute('width', Math.max(0, w - this.inset * 2));
    r.setAttribute('height', Math.max(0, h - this.inset * 2));
    return `url(#${id})`;
  },

  mountOneLine(text, d){
    while (text.firstChild) text.removeChild(text.firstChild);

    const tName = document.createElementNS('http://www.w3.org/2000/svg', 'tspan');
    tName.textContent = d.data.name || '';

    const tSep = document.createElementNS('http://www.w3.org/2000/svg', 'tspan');
    tSep.textContent = '  ';

    const tAvg = document.createElementNS('http://www.w3.org/2000/svg', 'tspan');
    tAvg.textContent = `整體：${displayPct(d.data.avg)}`;

    text.appendChild(tName);
    text.appendChild(tSep);
    text.appendChild(tAvg);
    text.dataset.mode = 'one';
  },

  mountTwoLines(text, d){
    while (text.firstChild) text.removeChild(text.firstChild);

    const tName = document.createElementNS('http://www.w3.org/2000/svg', 'tspan');
    tName.textContent = d.data.name || '';

    const tAvg = document.createElementNS('http://www.w3.org/2000/svg', 'tspan');
    tAvg.textContent = `整體：${displayPct(d.data.avg)}`;

    text.appendChild(tName);
    text.appendChild(tAvg);
    text.dataset.mode = 'two';
  },

  ellipsizeName(text, maxW){
    const tName = text.querySelector('tspan');
    if (!tName) return false;
    let nm = tName.textContent || '';
    if (nm.length === 0) return false;
    tName.textContent = nm.slice(0, -1) + '…';
    return true;
  },

  shortenAvg(text){
    const tsp = text.querySelectorAll('tspan');
    if (tsp.length === 0) return;

    const last = tsp[tsp.length - 1];
    const m = String(last.textContent || '').match(/([+\-]?[0-9]+(?:\.[0-9])?%)/);
    if (m) last.textContent = m[1];
  },

  fit(text, d, headerH){
    const wMaxFull = Math.max(0, d.x1 - d.x0) - this.inset * 2 - 2;
    const hMax = Math.max(0, headerH) - this.inset * 2 - 1;
    if (wMaxFull <= 0 || hMax <= 0) return;

    text.setAttribute('text-anchor', 'start');
    text.setAttribute('dominant-baseline', 'middle');
    text.setAttribute('x', d.x0 + this.inset + 4);
    text.setAttribute('y', d.y0 + headerH / 2);
    text.removeAttribute('lengthAdjust');
    text.removeAttribute('textLength');
    text.setAttribute('clip-path', this.ensureHeaderClip(text.ownerSVGElement, text.parentNode, d, headerH));

    this.mountOneLine(text, d);

    let f = Math.max(this.minFont, Math.floor(Math.min(Math.sqrt(Math.max(1, wMaxFull * hMax)) * this.k, hMax * 0.95)));
    let guard = 0;

    const loop = () => {
      if (++guard > 160) return;

      text.setAttribute('font-size', f);
      const mode = text.dataset.mode || 'one';
      const bb = text.getBBox();
      const sW = wMaxFull / Math.max(1, bb.width), sH = hMax / Math.max(1, bb.height);
      const s = Math.min(sW, sH, 1);
      const next = Math.max(this.minFont, Math.floor(f * s));

      if (next < f) {
        f = next;
        return loop();
      }

      if (sW < 1 && f <= this.minFont) {
        if (mode === 'one') {
          if (!this.ellipsizeName(text, wMaxFull)) {
            if (hMax >= this.minFont * 2 * this.lineHeight + 2) {
              this.mountTwoLines(text, d);
              return loop();
            }
          }
          return loop();
        } else {
          if (this.ellipsizeName(text, wMaxFull)) return loop();
        }
      }
      return;
    };

    loop();

    let bb = text.getBBox();
    if (bb.width > wMaxFull + 0.1) {
      this.shortenAvg(text);
      text.setAttribute('font-size', Math.max(this.minFont, parseInt(text.getAttribute('font-size') || this.minFont, 10) - 1));
      bb = text.getBBox();
    }

    if (bb.width > wMaxFull + 0.1) {
      text.setAttribute('lengthAdjust', 'spacingAndGlyphs');
      text.setAttribute('textLength', Math.max(1, Math.floor(wMaxFull)));
    }
  }
};


function renderTreemap(svgId, hintId, edges, codeField, month, metric, colorMode){
  const svgEl = document.getElementById(svgId);
  if (!svgEl) {
    console.warn(`[renderTreemap] 找不到 #${svgId}`);
    return;
  }

  const hint = document.getElementById(hintId) || null;
  const svg = d3.select(svgEl);
  svg.selectAll('*').remove();

  const wrap = svgEl.parentElement;
  if (!wrap) {
    console.warn(`[renderTreemap] #${svgId} 沒有 parentElement`);
    return;
  }

  const W = Math.max(320, (wrap.clientWidth || 0) - 16);
  const H = Math.max(320, parseInt(getComputedStyle(svgEl).height, 10) || 560);
  svg.attr('width', W).attr('height', H);

const groups = new Map();

for (const e of (edges || [])) {
  const keyRaw = normCode(e[codeField] || e['down'] || e['up']);
  if (isUSCode(keyRaw)) continue;

  const r = byCode.get(keyRaw);

  // 找不到 Revenue 資料者，不納入平均，也不放進熱力圖
  if (!r) continue;

  const v = getMetricValue(r, month, metric);

  // 平均值只排除無效值，不排除 0 或太小的數值
  // 這樣群組平均才會反映完整產業 / 概念股清單
  if (!Number.isFinite(v)) continue;

  const groupName = getTreemapGroupName(svgId, e, r);

  if (!groups.has(groupName)) {
    groups.set(groupName, {
      avgList: [],     // 完整有效清單：專門用來算平均
      renderList: []   // 顯示清單：專門用來畫熱力圖
    });
  }

  const codeVal =
    r['個股'] ??
    r['代號'] ??
    r['股票代碼'] ??
    r['股票代號'] ??
    r['公司代號'] ??
    r['證券代號'];

  const nameVal =
    r['名稱'] ??
    r['公司名稱'] ??
    r['證券名稱'];

const item = {
  code: codeVal,
  name: nameVal,
  raw: v,
  rel: groupName,
  row: r
};

  // 平均值：完整有效個股都納入
  groups.get(groupName).avgList.push(item);

  // 熱力圖：仍維持原本規則，0、接近 0、無效值不顯示
  if (!shouldSkipTreemapValue(v)) {
    groups.get(groupName).renderList.push(item);
  }
}

  if (groups.size === 0) {
    if (hint) hint.textContent = '查無關聯個股';
    return;
  }

  const EPS = 0.01;
  const allSummaries = [];

for (const [rel, groupObj] of groups) {
  const avgList = groupObj.avgList || [];
  const renderList = groupObj.renderList || [];

// 群組表現：
// 左邊上游熱力圖改用「整體營收加總」計算
// MoM = 本月營收加總 / 上月營收加總 - 1
// YoY = 本月營收加總 / 去年同月營收加總 - 1
//
// 右邊下游概念股如果你也想套用同樣邏輯，可以拿掉 svgId === 'upTreemap' 的限制

  const groupRows = avgList
    .map(d => d.row)
    .filter(Boolean);

 let avg = getGroupAggregateRevenuePerformance(groupRows, month, metric);

if (!Number.isFinite(avg)) {
  avg = 0;
}

  // 熱力圖只使用實際可顯示清單
  const baseValues = renderList.map(s => {
    const base = getTreemapLeafBase(s.raw, svgId);
    return { s, base };
  });

  const baseSum = d3.sum(baseValues, d => d.base) || EPS;

  allSummaries.push({
    rel,
    list: renderList,      // 畫圖用
    avgList,               // 平均用完整清單
    avg,
    baseValues,
    baseSum
  });
}

  const groupSummaries = selectTreemapGroups(svgId, allSummaries);

  if (groupSummaries.length === 0) {
    if (hint) hint.textContent = '沒有符合條件的群組';
    return;
  }

let groupWeights = new Map();

/*
  群組面積權重：
  - 上游 / 相同產業股：維持原本 RANK 排名邏輯，避免單一類股過度放大
  - 下游 / 概念股：改用平均營收表現決定面積，讓平均表現高的概念股群組更明顯
*/
if (svgId === 'downTreemap') {
  for (const g of groupSummaries) {
    const avg = Number.isFinite(g.avg) ? g.avg : 0;

    // 加 10 是避免低平均或小負值群組被壓到太小
    // Math.max(0, avg) 代表負值概念股不會因為絕對值大而被放大
    const score = Math.max(0, avg) + 10;

    groupWeights.set(g.rel, Math.max(EPS, score));
  }
} else if (GROUP_WEIGHT_MODE === 'AVG') {
  const minAvg = d3.min(groupSummaries.map(d => Number.isFinite(d.avg) ? d.avg : 0));

  for (const g of groupSummaries) {
    const a = Number.isFinite(g.avg) ? g.avg : minAvg;
    groupWeights.set(g.rel, Math.max(EPS, (a - minAvg + EPS)));
  }
} else {
  const sorted = [...groupSummaries].sort((a, b) => {
    const av = Number.isFinite(a.avg) ? a.avg : -Infinity;
    const bv = Number.isFinite(b.avg) ? b.avg : -Infinity;
    return av - bv;
  });

  const n = Math.max(1, sorted.length - 1);

  sorted.forEach((g, i) => {
    const t = i / n;
    const w = RANK_WEIGHT_MIN + t * (RANK_WEIGHT_MAX - RANK_WEIGHT_MIN);
    groupWeights.set(g.rel, w);
  });
}

let children = [];

for (const g of groupSummaries) {
  const gw = groupWeights.get(g.rel) || 1;

  let baseValuesForRender = g.baseValues;

  // 右邊概念股：保留表現較佳的前幾檔
  // 但後面一定要用「實際渲染清單」重新計算分母，避免群組被全部名單稀釋
  if (svgId === 'downTreemap') {
    baseValuesForRender = [...g.baseValues]
      .sort((a, b) => {
        const av = Number.isFinite(a.s.raw) ? a.s.raw : -Infinity;
        const bv = Number.isFinite(b.s.raw) ? b.s.raw : -Infinity;

        if (bv !== av) return bv - av;

        return String(a.s.code || '').localeCompare(
          String(b.s.code || ''),
          'zh-Hant',
          { numeric: true }
        );
      })
      .slice(0, DOWN_GROUP_STOCK_KEEP_MAX);
  }

  // 重要修正：
  // 只用實際要畫出的個股計算 baseSum
  // 不再用 g.baseSum，避免輝達這種名單很多的概念股被稀釋
  const renderBaseSum = d3.sum(baseValuesForRender, d => d.base) || EPS;
  const scale = gw / renderBaseSum;

const kids = baseValuesForRender.map(({s, base}) => ({
  name: s.name || '',
  code: s.code,
  raw: s.raw,
  rel: s.rel || g.rel,
  value: base * scale
}));

// 避免空群組進入 D3 layout，造成後續判斷異常
if (!kids.length) continue;

children.push({
  name: g.rel,
  avg: g.avg,
  targetValue: gw,
  children: kids
});
}

  let root = d3.hierarchy({ children })
    .sum(d => d.value)
    .sort(treemapSortByPerformance);

  d3.treemap()
    .tile(d3.treemapSquarify.ratio(1.15))
    .size([W, H])
    .paddingOuter(8)
    .paddingInner(3)
    .paddingTop(HEADER_H)(root);

const filteredChildren = (root.children || []).map(parent => {
  const leaves = parent.children || [];

  let keptLeaves = leaves.filter(leaf => {
    const w = Math.max(0, leaf.x1 - leaf.x0);
    const h = Math.max(0, leaf.y1 - leaf.y0);
    const area = w * h;

    // 左邊相同產業股維持全部顯示
    if (svgId === 'upTreemap') return true;

    // 右邊概念股改成自適應門檻
    // 畫面寬度較小時，自動降低過濾條件
    const adaptiveMinW = Math.max(
      DOWN_MIN_RENDER_W_FLOOR,
      Math.min(MIN_RENDER_W, W * 0.08)
    );

    const adaptiveMinH = Math.max(
      DOWN_MIN_RENDER_H_FLOOR,
      Math.min(MIN_RENDER_H, H * 0.035)
    );

    const adaptiveMinArea = Math.max(
      DOWN_MIN_RENDER_AREA_FLOOR,
      Math.min(MIN_RENDER_AREA, W * H * 0.00045)
    );

    return w >= adaptiveMinW && h >= adaptiveMinH && area >= adaptiveMinArea;
  });

  // ===== 關鍵保底修正 =====
  // 如果右側概念股在 100% 畫面下全部被濾掉，
  // 不要讓整個群組消失，至少保留表現最好的前幾檔。
  if (!keptLeaves.length && svgId === 'downTreemap') {
    keptLeaves = [...leaves]
      .sort((a, b) => {
        const av = Number.isFinite(a.data?.raw) ? a.data.raw : -Infinity;
        const bv = Number.isFinite(b.data?.raw) ? b.data.raw : -Infinity;

        if (bv !== av) return bv - av;

        return String(a.data?.code || '').localeCompare(
          String(b.data?.code || ''),
          'zh-Hant',
          { numeric: true }
        );
      })
      .slice(0, DOWN_FALLBACK_KEEP_PER_GROUP);
  }

  if (!keptLeaves.length) return null;

  // 個股被濾掉後，重新把保留下來的個股 value 放大回原本群組權重
  const targetValue = Number.isFinite(parent.data.targetValue)
    ? parent.data.targetValue
    : d3.sum(keptLeaves, leaf => leaf.data.value);

  const keptSum = d3.sum(keptLeaves, leaf => leaf.data.value) || EPS;
  const rescale = targetValue / keptSum;

  const childrenRescaled = keptLeaves.map(leaf => ({
    ...leaf.data,
    value: leaf.data.value * rescale
  }));

  return {
    name: parent.data.name,
    avg: parent.data.avg,
    targetValue,
    children: childrenRescaled
  };
}).filter(Boolean);


if (filteredChildren.length === 0) {
  if (hint) {
    hint.textContent = svgId === 'downTreemap'
      ? '概念股資料存在，但目前畫面寬度不足，已無可顯示方塊'
      : '此區個股方塊過小，已自動省略';
  }
  return;
}

  root = d3.hierarchy({ children: filteredChildren })
    .sum(d => d.value)
    .sort(treemapSortByPerformance);

  d3.treemap()
    .tile(d3.treemapSquarify.ratio(1.15))
    .size([W, H])
    .paddingOuter(8)
    .paddingInner(3)
    .paddingTop(HEADER_H)(root);

  const g = svg.append('g');

  const parents = g.selectAll('g.parent')
    .data(root.children || [])
    .enter()
    .append('g')
    .attr('class', 'parent');

  parents.append('rect')
    .attr('class', 'group-bg')
    .attr('x', d => d.x0)
    .attr('y', d => d.y0)
    .attr('width', d => Math.max(0, d.x1 - d.x0))
    .attr('height', d => Math.max(0, d.y1 - d.y0))
    .attr('fill', d => colorFor(d.data.avg, colorMode));

  parents.append('rect')
    .attr('class', 'group-border')
    .attr('x', d => d.x0)
    .attr('y', d => d.y0)
    .attr('width', d => Math.max(0, d.x1 - d.x0))
    .attr('height', d => Math.max(0, d.y1 - d.y0));

  const titles = parents.append('text')
    .attr('class', 'node-title')
    .attr('fill', '#fff')
    .style('paint-order', 'stroke')
    .style('stroke', 'rgba(0,0,0,0.35)')
    .style('stroke-width', '2px');

  titles.each(function(d){ GroupTitleFit.fit(this, d, HEADER_H); });

  const node = g.selectAll('g.node')
    .data(root.leaves())
    .enter()
    .append('g')
    .attr('class', 'node')
    .attr('transform', d => `translate(${d.x0},${d.y0})`);

  node.append('rect')
    .attr('class', 'node-rect')
    .attr('width', d => Math.max(0, d.x1 - d.x0))
    .attr('height', d => Math.max(0, d.y1 - d.y0))
    .attr('fill', d => colorFor(d.data.raw, colorMode));

  const labels = node.append('text')
    .attr('class', 'node-label')
    .attr('fill', '#fff')
    .style('paint-order', 'stroke')
    .style('stroke', 'rgba(0,0,0,0.35)')
    .style('stroke-width', '2px')
    .style('text-rendering', 'geometricPrecision');

  labels.each(function(d){
    const code = `${d.data.code || ''}`.trim();
    const name = `${d.data.name || ''}`.trim();
    const pct  = displayPct(d.data.raw);
    const rel  = `${d.data.rel || ''}`.trim();

    this.dataset.code = code;
    this.dataset.name = name;
    this.dataset.pct = pct;

    const lines = [code, name, pct].filter(Boolean);
    lines.forEach(line => {
      const t = document.createElementNS('http://www.w3.org/2000/svg', 'tspan');
      t.textContent = line;
      this.appendChild(t);
    });

    const title = document.createElementNS('http://www.w3.org/2000/svg', 'title');
    title.textContent = `${code} ${name}\n${rel}\n${month.slice(0,4)}/${month.slice(4,6)} ${metric}: ${pct}`;
    this.appendChild(title);
  });

  if (ENABLE_NODE_CLICK) {
    node
      .style('cursor', 'pointer')
      .on('click', function(event, d){
        const code = `${d.data.code || ''}`.trim();
        if (!code) return;

        const input = document.querySelector('#stockInput');
        if (input) input.value = code;

        triggerUnifiedQuery();
      });
  }

  requestAnimationFrame(() => {
    node.each(function(d){
      const w = Math.max(0, d.x1 - d.x0);
      const h = Math.max(0, d.y1 - d.y0);
      const textEl = this.querySelector('text');
      if (!textEl) return;

      try {
        LabelFit.fitBlock(textEl, w, h);
        LabelFit.ensureClip(this, w, h);
      } catch (err) {
        console.error(`[LabelFit ${svgId}]`, err, d);
      }
    });

    parents.select('text').each(function(d){
      try {
        GroupTitleFit.fit(this, d, HEADER_H);
      } catch (err) {
        console.error(`[GroupTitleFit ${svgId}]`, err, d);
      }
    });
  });

  const onResize = () => {
    parents.select('text').each(function(d){
      try {
        GroupTitleFit.fit(this, d, HEADER_H);
      } catch (err) {
        console.error(`[resize title fit ${svgId}]`, err, d);
      }
    });
  };

  window.addEventListener('resize', onResize, { passive:true });
}

function toNum(v){
  if (v == null || v === '') return null;
  if (typeof v === 'string') {
    v = v.replace(/[%％,\s]/g, '').trim();
  }
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

function getLatestMonthLabel(){
  const m = months && months.length ? months[0] : '';
  if (!m) return '最新月份';

  const year = m.slice(0, 4);
  const month = String(Number(m.slice(4, 6))); // 去掉前導 0，例如 03 -> 3

  return `${year}年${month}月`;
}


function formatMonthLabelForConceptTable(month){
  if (!month || String(month).length !== 6) return '';

  const year = String(month).slice(0, 4);
  const mm = String(month).slice(4, 6);

  return `${year}年${mm}月`;
}





function getStockPageUrl(code){
  const c = encodeURIComponent(normCode(code));
  return `https://www.fbs.com.tw/MKT/Index?name=%EF%BC%AA%E7%B7%9A%E5%9C%96&stock=${c}`;
}



function ensureNewHighTableStyles(){
  if (document.getElementById('newHighTableInlineStyle')) return;

  const style = document.createElement('style');
  style.id = 'newHighTableInlineStyle';
  style.textContent = `
    #newHighTableWrap .new-high-table {
      width: 100% !important;
      border-collapse: collapse !important;
      table-layout: fixed !important;
    }

    #newHighTableWrap .new-high-table col.col-code {
      width: 18% !important;
    }
    #newHighTableWrap .new-high-table col.col-name {
      width: 25% !important;
    }
    #newHighTableWrap .new-high-table col.col-mom {
      width: 18% !important;
    }
    #newHighTableWrap .new-high-table col.col-yoy {
      width: 18% !important;
    }


    #newHighTableWrap .new-high-table thead th,
    #newHighTableWrap .new-high-table tbody td {
      padding: 10px 14px !important;
      vertical-align: middle !important;
      box-sizing: border-box !important;
    }

    /* 表頭 */
    #newHighTableWrap .new-high-table thead th {
      color: #dbeafe !important;
      font-weight: 600 !important;
      text-align: left !important;
      white-space: nowrap !important;
      border-bottom: 1px solid rgba(255,255,255,0.10) !important;
    }

    /* 數值欄位：表頭與內容統一靠右 */
    #newHighTableWrap .new-high-table thead th.th-num,
    #newHighTableWrap .new-high-table tbody td.num {
      text-align: right !important;
      font-variant-numeric: tabular-nums !important;
    }

    /* 文字欄位：統一靠左 */
    #newHighTableWrap .new-high-table tbody td.code,
    #newHighTableWrap .new-high-table tbody td.name {
      text-align: left !important;
    }

    /* 每個欄位中間加淡白線 */
    #newHighTableWrap .new-high-table thead th + th,
    #newHighTableWrap .new-high-table tbody td + td {
      border-left: 1px solid rgba(255,255,255,0.07) !important;
    }

    /* 每列底部淡線 */
    #newHighTableWrap .new-high-table tbody tr.stock-row td {
      border-bottom: 1px dashed rgba(255,255,255,0.05) !important;
    }

    /* 產業標題列 */
    #newHighTableWrap .new-high-table tbody tr.group-row td {
      background: rgba(37, 99, 235, 0.12) !important;
      color: #eaf2ff !important;
      font-weight: 600 !important;
      border-top: 1px solid rgba(255,255,255,0.06) !important;
      border-bottom: 1px solid rgba(255,255,255,0.08) !important;
      padding: 9px 14px !important;
    }

    /* 股票可點擊文字 */
    #newHighTableWrap .stock-link {
      display: inline-block !important;
      color: #f8fafc !important;
      cursor: pointer !important;
      text-decoration: none !important;
      line-height: 1.2 !important;
    }

    #newHighTableWrap .stock-link:hover {
      color: #93c5fd !important;
      text-decoration: underline !important;
    }

    /* 群組 +/- 按鈕 */
    #newHighTableWrap .group-toggle {
      display: inline-flex !important;
      width: 14px !important;
      justify-content: center !important;
      align-items: center !important;
      margin-right: 8px !important;
      cursor: pointer !important;
      color: #cfe3ff !important;
      user-select: none !important;
      font-weight: 700 !important;
    }
  `;

  document.head.appendChild(style);
}


function extractNewHighRecords(){
  if (!Array.isArray(newHighSheetRows) || newHighSheetRows.length === 0) return [];

  const out = [];

  for (const row of newHighSheetRows) {
    // 依你指定欄位：
    // A=0 股票代號
    // B=1 股票名稱
    // F=5 最新月 MoM
    // G=6 最新月 YoY
    // N=13 是否創新高（A）
    // O=14 產業類別
    const code = normCode(row[0]);
    const name = normText(row[1]);
    const mom  = toNum(row[5]);
    const yoy  = toNum(row[6]);
    const flag = normText(row[13]).toUpperCase();
    const industry = normText(row[14]) || '未分類';

    // 略過前面說明列 / 標題列
    if (!code || !name) continue;
    if (code === '股票代號' || name === '股票名稱') continue;

    // 只保留 N欄 = A
    if (flag !== 'A') continue;

    out.push({
      code,
      name,
      mom,
      yoy,
      industry
    });
  }

  return out;
}

function groupAndSortNewHighRecords(records){
  const groups = new Map();

  for (const r of records) {
    if (!groups.has(r.industry)) groups.set(r.industry, []);
    groups.get(r.industry).push(r);
  }

  const result = [...groups.entries()].map(([industry, list]) => {
    // 同產業內：改成依個股代號排序
    list.sort((a, b) => {
      return String(a.code).localeCompare(String(b.code), 'zh-Hant', { numeric: true });
    });

    return { industry, list };
  });

  // 產業排序：
  // 1. 先把「未分類-傳產 / 未分類-電子」放最後
  // 2. 其他依家數多到少
  // 3. 同家數再依產業名稱排序
  result.sort((a, b) => {
    const tail = ['未分類-傳產', '未分類-電子'];

    const aTail = tail.includes(a.industry);
    const bTail = tail.includes(b.industry);

    if (aTail && !bTail) return 1;
    if (!aTail && bTail) return -1;

    if (b.list.length !== a.list.length) {
      return b.list.length - a.list.length;
    }

    return a.industry.localeCompare(b.industry, 'zh-Hant');
  });

  return result;
}

function renderNewHighSummary(){
  const host = document.getElementById('newHighTableWrap');
  const titleEl = document.getElementById('newHighTitle');
  const metaEl = document.getElementById('newHighMeta');

  if (!host) return;

  // 先確保樣式已經注入
  ensureNewHighTableStyles();

  const records = extractNewHighRecords();
  const groups = groupAndSortNewHighRecords(records);
  const latestMonthLabel = getLatestMonthLabel();

  if (titleEl) {
    titleEl.textContent = `${latestMonthLabel} 營收創新高個股彙整`;
  }

  if (!records.length) {
    if (metaEl) metaEl.textContent = '沒有符合條件的資料';
    host.innerHTML = `<div class="new-high-empty">最新月份沒有營收創新高個股資料。</div>`;
    return;
  }

  if (metaEl) {
    metaEl.textContent = `共 ${records.length} 檔｜${groups.length} 個產業類別`;
  }

  let visibleStockCount = 0;
  const bodyRows = [];

  for (const g of groups) {
    const groupStartVisibleCount = visibleStockCount;
    const groupInitiallyExpanded = groupStartVisibleCount < NEWHIGH_COLLAPSE_AFTER;

  const stockRowsHtml = g.list.map(r => {
    visibleStockCount += 1;
    const isExtraRow = visibleStockCount > NEWHIGH_COLLAPSE_AFTER;
    const stockUrl = getStockPageUrl(r.code);

    return `
      <tr class="stock-row ${isExtraRow ? 'extra-row' : ''}" data-industry="${safe(g.industry)}">
        <td class="code">
          <a
            class="stock-link"
            data-code="${safe(r.code)}"
            href="${safe(stockUrl)}"
            target="_blank"
            rel="noopener noreferrer"
          >
            ${safe(r.code)}
          </a>
        </td>
        <td class="name">
          <a
            class="stock-link"
            data-code="${safe(r.code)}"
            href="${safe(stockUrl)}"
            target="_blank"
            rel="noopener noreferrer"
          >
            ${safe(r.name)}
          </a>
        </td>
        <td class="num">${displayPct(r.mom)}</td>
        <td class="num">${displayPct(r.yoy)}</td>
      </tr>
    `;
  }).join('');

    const groupHeaderHtml = `
      <tr class="group-row" data-industry="${safe(g.industry)}">
        <td colspan="4">
          <span
            class="group-toggle"
            data-industry="${safe(g.industry)}"
            data-expanded="${groupInitiallyExpanded ? 'true' : 'false'}"
          >${groupInitiallyExpanded ? '－' : '＋'}</span>
          ${safe(g.industry)}（${g.list.length} 檔）
        </td>
      </tr>
    `;

    bodyRows.push(groupHeaderHtml + stockRowsHtml);
  }

  host.innerHTML = `
    <div class="new-high-table-wrap">
      <table class="new-high-table">
        <colgroup>
          <col class="col-code">
          <col class="col-name">
          <col class="col-mom">
          <col class="col-yoy">
        </colgroup>
        <thead>
          <tr>
            <th class="th-code">個股代號</th>
            <th class="th-name">個股名稱</th>
            <th class="th-num">MoM</th>
            <th class="th-num">YoY</th>
          </tr>
        </thead>
        <tbody>
          ${bodyRows.join('')}
        </tbody>
      </table>
    </div>
  `;

  // 預設先把第 15 檔之後的資料藏起來
  host.querySelectorAll('.extra-row').forEach(row => {
    row.style.display = 'none';
  });



  // 每個產業自己的展開 / 收合
  host.querySelectorAll('.group-toggle').forEach(btn => {
    btn.addEventListener('click', () => {
      const industry = btn.dataset.industry;
      const expanded = btn.dataset.expanded === 'true';

      const rows = host.querySelectorAll(
        `.stock-row[data-industry="${CSS.escape(industry)}"]`
      );

      const nextExpanded = !expanded;
      btn.dataset.expanded = nextExpanded ? 'true' : 'false';
      btn.textContent = nextExpanded ? '－' : '＋';

      rows.forEach(row => {
        row.style.display = nextExpanded ? '' : 'none';
      });
    });
  });
}
