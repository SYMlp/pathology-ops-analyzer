
const ML = ['1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月'];

function fmtMoney(num) { return fmtNum(num); }
function fmtNum(num) { 
  if (num == null) return '--'; 
  if (typeof num === 'string') { num = parseFloat(num.replace(/,/g, '')); if (isNaN(num)) return '--'; }
  const valstr = num.toLocaleString('zh-CN', { minimumFractionDigits: 0, maximumFractionDigits: 2 });
  if (Math.abs(num) >= 10000) return `<span class="val-trunc" title="${valstr}">` + (num/10000).toFixed(1) + '<small>万</small></span>';
  if (Math.abs(num) >= 1000) return `<span class="val-trunc" title="${valstr}">` + (num/1000).toFixed(1) + '<small>千</small></span>';
  return Math.round(num).toString();
}
function fmtChange(num) { if (num == null) return '--'; return (num > 0 ? '+' : '') + (num * 100).toFixed(1) + '%'; }
function fmtPct(num) { return num == null ? '--' : (num * 100).toFixed(1) + '%'; }
function esc(str) { return (str || '').replace(/[&<>'"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' })[c]); }

function changeClass(val, invert = false) {
  if (val == null || val === 0) return '';
  const positive = val > 0;
  return (invert ? !positive : positive) ? 'val-pos' : 'val-neg';
}

function getActiveMonths(trendObj) {
  if (!trendObj) return [];
  return Object.keys(trendObj).map(Number).sort((a, b) => a - b).filter(m => trendObj[m] != null);
}

function statusOf(val, goodThresh, badThresh, invert = false) {
  if (val == null) return 'na';
  if (!invert) {
    if (val <= goodThresh) return 'good';
    if (val > badThresh) return 'bad';
    return 'warn';
  } else {
    if (val >= goodThresh) return 'good';
    if (val < badThresh) return 'bad';
    return 'warn';
  }
}

function reportCSS() { return `<style>:root{
  --bg:#f9f8f4;--bg-paper:#ffffff;--surface:#ffffff;--surface-2:#f4f1ea;--border:#e6dfd1;--ink:#423931;--muted:#6e655b;--faint:#a1968b;
  --teal:#3a6e57;--teal-dim:#eef4f1;--coral:#bc5e4c;--coral-dim:#fbeee9;--indigo:#8b6b52;--indigo-dim:#f5efe9;
  --radius:8px;--radius-sm:6px;
  --shadow:0 4px 16px rgba(100,85,70,.05), 0 1px 3px rgba(100,85,70,.03);
  --font-sans:'Inter',system-ui,-apple-system,sans-serif;--font-serif:'Noto Serif SC','SimSun',serif;
}
*{margin:0;padding:0;box-sizing:border-box}
body{
  font-family:var(--font-sans);background:var(--bg);color:var(--ink);line-height:1.5;font-size:13px;
  -webkit-font-smoothing:antialiased;
}
.report{max-width:1080px;margin:0 auto;padding:24px 32px 32px}
small{font-size:.65em;font-weight:500;color:var(--faint)}

.report-header{
  position:relative;padding:20px 24px;background:#ffffff;
  color:var(--ink);border-radius:var(--radius);margin-bottom:14px;overflow:hidden;border:1px solid var(--border);
  box-shadow:var(--shadow);
}
.header-bg{display:none}
.header-content{position:relative;z-index:1}
.header-badge{display:inline-block;padding:3px 10px;background:var(--surface-2);color:var(--indigo);border-radius:4px;font-size:10px;font-weight:700;letter-spacing:.08em;margin-bottom:8px;border:1px solid var(--border)}
.report-header h1{font-family:var(--font-serif);font-size:clamp(1.35rem,3.5vw,1.65rem);font-weight:700;letter-spacing:.02em;margin-bottom:8px}
.header-meta{display:flex;flex-wrap:wrap;align-items:center;gap:8px 14px;font-size:13px;color:var(--muted)}
.header-tag{padding:3px 10px;background:var(--surface-2);border-radius:6px;font-weight:500;color:var(--ink);border:1px solid var(--border)}
.header-divider{width:1px;height:14px;background:var(--border)}
.header-actions{position:absolute;top:20px;right:24px;z-index:1}
.btn-action{display:inline-flex;align-items:center;gap:6px;padding:8px 14px;background:var(--surface-2);color:var(--ink);border:1px solid var(--border);border-radius:6px;cursor:pointer;font-size:12px;font-weight:600;transition:background .2s}
.btn-action:hover{background:#e6dfd1}

.ai-summary{background:#fffcf8;border:1px solid var(--border);border-left:4px solid var(--indigo);border-radius:var(--radius-sm);padding:14px 16px;margin-bottom:14px;position:relative;display:flex;gap:12px;align-items:flex-start}
.ai-summary-icon{display:none}
.ai-summary-text{flex:1;font-size:13px;line-height:1.5;color:var(--ink)}
.ai-summary-text strong{color:var(--indigo);font-weight:700}

.exec-summary{background:var(--surface);border-radius:var(--radius);padding:14px 18px;margin-bottom:14px;border:1px solid var(--border);box-shadow:var(--shadow)}
.exec-summary-head{display:flex;align-items:baseline;justify-content:space-between;gap:8px;margin-bottom:14px;padding-bottom:10px;border-bottom:1px solid var(--border)}
.exec-title{font-family:var(--font-serif);font-size:1.15rem;font-weight:700;color:var(--ink)}
.exec-sub{font-size:10px;color:var(--faint);text-transform:uppercase;letter-spacing:.04em}

.kpi-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:12px}
.kpi-card{
  background:var(--surface);border-radius:var(--radius-sm);padding:12px 14px;
  border:1px solid var(--border);position:relative;overflow:hidden;
  display:grid;grid-template-areas:"icon label""icon value""icon sub""gauge gauge";
  grid-template-columns:36px 1fr;gap:2px 10px;align-items:center;
  box-shadow:0 1px 2px rgba(90,75,60,.02);
}
.kpi-card:hover{border-color:#d4c8b6;box-shadow:var(--shadow)}
.kpi-icon{grid-area:icon;margin:0;width:36px;height:36px;border-radius:8px;display:flex;align-items:center;justify-content:center;color:var(--indigo);background:var(--indigo-dim);align-self:start;margin-top:2px}
.kpi-icon svg{width:20px;height:20px}
.kpi-content{display:contents}
.kpi-icon-income{color:var(--teal);background:var(--teal-dim)}.kpi-icon-expense{color:var(--coral);background:var(--coral-dim)}
.kpi-icon-surplus{color:var(--indigo);background:var(--indigo-dim)}.kpi-icon-ratio{color:#b46c24;background:#f8f2e6}
.kpi-icon-workload{color:#365b98;background:#ebf0f7}.kpi-icon-percapita{color:#6c4b8b;background:#f2ecf6}
.kpi-icon-satisfaction{color:#aa721f;background:#fbf3e5}.kpi-icon-team{color:var(--muted);background:var(--surface-2)}
.kpi-label{grid-area:label;align-self:end;font-size:10px;color:var(--muted);font-weight:600;letter-spacing:.02em;margin-bottom:-2px;white-space:nowrap}
.kpi-value{grid-area:value;font-size:clamp(1.1rem,2vw,1.3rem);font-weight:700;color:var(--ink);font-variant-numeric:tabular-nums;line-height:1}
.kpi-sub{grid-area:sub;align-self:start;font-size:9px;font-weight:600;margin-top:2px}
.val-pos{color:var(--teal)}.val-neg{color:var(--coral)}
.gauge-bar{grid-area:gauge;height:4px;background:var(--surface-2);border-radius:2px;margin-top:6px;position:relative}
.gauge-fill{height:100%;border-radius:2px;transition:width .5s ease;background:var(--teal)}
.gauge-marks{position:absolute;inset:0}.gauge-marks span{position:absolute;top:-2px;width:1px;height:8px;background:var(--faint);opacity:.5}
.gauge-marks .mark-danger{background:var(--coral);opacity:.8}

.report-body{display:flex;flex-direction:column;gap:12px}

.section{background:var(--surface);border-radius:var(--radius);padding:12px 14px;border:1px solid var(--border);box-shadow:var(--shadow)}
.section-compact{padding:10px 12px}
.section-header{margin-bottom:8px}
.section-header-inline{display:flex;flex-wrap:wrap;align-items:center;gap:8px 12px;margin-bottom:8px;padding-bottom:6px;border-bottom:1px solid var(--border)}
.section-header-inline h2{font-size:0.95rem;font-weight:700;color:var(--ink);margin:0}
.section-desc{font-size:9px;color:var(--faint);text-transform:uppercase;letter-spacing:.05em}
.section-strip{font-size:11px;color:var(--muted);margin-left:auto;flex:1 1 auto;text-align:right}

.grid-2{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:8px}
.grid-tight{gap:8px;margin-bottom:8px}
.chart-card{background:var(--bg-paper);border-radius:var(--radius-sm);padding:8px 10px;border:1px solid var(--border)}
.chart-card h3{font-size:11px;color:var(--muted);font-weight:600;margin-bottom:6px}

.efficiency-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:6px}
.eff-card{background:var(--bg-paper);border-radius:var(--radius-sm);padding:8px 10px;border:1px solid var(--border);border-left:3px solid transparent}
.eff-good{border-left-color:var(--teal)}.eff-warn{border-left-color:#d97706}.eff-bad{border-left-color:var(--coral)}.eff-info{border-left-color:var(--indigo)}.eff-na{border-left-color:var(--faint)}
.eff-label{font-size:10px;color:var(--faint);font-weight:600;margin-bottom:2px}
.eff-value{font-size:0.95rem;font-weight:700;font-variant-numeric:tabular-nums}
.eff-bench{font-size:9px;color:var(--faint);margin-top:2px}

.stat-panel{background:var(--bg-paper);border-radius:var(--radius-sm);padding:8px 10px;border:1px solid var(--border)}
.stat-panel-dense h3{font-size:11px;color:var(--muted);font-weight:600;margin-bottom:4px}
.stat-row{display:flex;justify-content:space-between;align-items:center;padding:4px 0;border-bottom:1px solid var(--border);font-size:11px}
.stat-row.sub{padding:2px 0;border-bottom:1px dashed var(--border);font-size:10px}
.stat-row.sub span:first-child{color:var(--muted)}
.stat-big{font-size:0.95rem;font-weight:700;font-variant-numeric:tabular-nums}
.stat-divider{height:1px;background:var(--border);margin:4px 0}

.mini-cards{display:flex;gap:8px;margin-bottom:10px;flex-wrap:wrap}
.mini-cards-row{margin-bottom:8px}
.mini-card{flex:1;min-width:100px;background:var(--bg-paper);border-radius:var(--radius-sm);padding:8px 10px;border:1px solid var(--border);display:flex;flex-direction:column;gap:2px}
.mini-card.sm{padding:6px 8px;min-width:88px}
.mini-label{font-size:10px;color:var(--faint);font-weight:600;text-transform:uppercase;letter-spacing:.04em}
.mini-value{font-size:1rem;font-weight:700;font-variant-numeric:tabular-nums}
.mini-sub{font-size:10px;color:var(--faint)}

.ranking-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:8px}
.ranking-grid-dense{gap:6px}
.rank-card{background:var(--bg-paper);border-radius:var(--radius-sm);padding:10px 12px;text-align:center;position:relative;border:1px solid var(--border)}
.rank-top{background:linear-gradient(135deg,#eff6ff,#e0e7ff);border-color:#c7d2fe}
.rank-label{font-size:10px;color:var(--faint);font-weight:600;margin-bottom:4px}
.rank-value{font-size:1.05rem;font-weight:700;font-variant-numeric:tabular-nums}
.rank-badge{position:absolute;top:8px;right:8px;width:24px;height:24px;border-radius:50%;background:var(--indigo);color:#fff;font-size:10px;font-weight:700;display:flex;align-items:center;justify-content:center}

.report-fold{border:1px solid var(--border);border-radius:var(--radius-sm);background:var(--bg-paper);margin-top:8px;overflow:hidden}
.report-fold>.fold-summary{
  list-style:none;cursor:pointer;display:flex;align-items:center;gap:10px;padding:10px 12px;font-size:13px;font-weight:600;color:var(--ink);
  background:linear-gradient(90deg,rgba(61,79,124,.06),transparent);user-select:none;
}
.report-fold>.fold-summary::-webkit-details-marker{display:none}
.fold-chevron{flex-shrink:0;width:8px;height:8px;border-right:2px solid var(--muted);border-bottom:2px solid var(--muted);transform:rotate(-45deg);transition:transform .2s;margin-top:-2px}
.report-fold[open]>.fold-summary .fold-chevron{transform:rotate(45deg);margin-top:2px}
.fold-title-text{font-family:var(--font-serif);font-size:13px}
.fold-hint{margin-left:auto;font-size:11px;font-weight:500;color:var(--faint);text-align:right;max-width:55%}
.fold-body{padding:0 12px 12px}
.appendix-wrap .report-fold{margin-top:0}

.table-wrap{margin-top:0}
.table-wrap.tight{margin-top:10px}
.table-wrap.tight:first-child{margin-top:0}
.table-inline-title{font-size:11px;color:var(--muted);font-weight:700;text-transform:uppercase;letter-spacing:.06em;margin-bottom:6px}
.table-scroller{width:100%;overflow-x:auto;-webkit-overflow-scrolling:touch;padding-bottom:4px;margin-bottom:-4px}
.data-table{width:max-content;min-width:100%;border-collapse:separate;border-spacing:0;font-size:12px;border:1px solid var(--border);border-radius:var(--radius-sm)}
.data-table-dense th,.data-table-dense td{padding:6px 8px}
.data-table th{background:var(--surface-2);color:var(--muted);padding:8px 10px;font-weight:700;font-size:10px;text-transform:uppercase;letter-spacing:.04em;border-bottom:1px solid var(--border)}
.data-table td{padding:7px 10px;border-bottom:1px solid var(--border);background:var(--surface)}
.data-table tr:last-child td{border-bottom:none}
.data-table tr:hover td{background:rgba(61,79,124,.05)}
.data-table th.st-col, .data-table td.st-col{position:sticky;z-index:2;box-shadow:inset -1px 0 0 var(--border)}
.data-table th.st-col{z-index:3;background:var(--surface-2)}
.data-table .st-1{left:0;width:125px;min-width:125px;max-width:125px;white-space:normal;word-break:break-all}
.data-table .st-2{left:125px;width:75px;min-width:75px;max-width:75px}
.data-table .st-3{left:200px;width:75px;min-width:75px;max-width:75px}
.data-table .st-divider{box-shadow:inset -2px 0 0 var(--border)}
.data-table .st-r1{right:75px;width:75px;min-width:75px;max-width:75px;box-shadow:inset 2px 0 0 var(--border);border-left:none}
.data-table .st-r2{right:0;width:75px;min-width:75px;max-width:75px;box-shadow:none;border-left:none}
.data-table tr:hover td.st-col{background:#f4f6f9}
.data-table .col-highlight{background:var(--indigo-dim);color:var(--indigo);font-weight:700}
.data-table tr:hover td.col-highlight{background:#dce1f0}
.val-trunc{cursor:help;border-bottom:1px dashed var(--faint);padding-bottom:1px;position:relative;z-index:1}
.val-trunc small{font-size:0.8em;margin-left:1px}
.data-table .td-name{font-weight:600}
.data-table .td-ellipsis{max-width:220px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.data-table .r{text-align:right}
.data-table .c{text-align:center}
.data-table .num{font-variant-numeric:tabular-nums;font-family:ui-monospace,'IBM Plex Sans',monospace}
.data-table tr.row-not-parsed td{color:var(--faint)}

.tip-icon{display:inline-flex;align-items:center;justify-content:center;width:12px;height:12px;border-radius:50%;background:#94a3b8;color:#fff;font-size:9px;font-weight:700;cursor:pointer;margin-left:4px;vertical-align:middle;position:relative;top:-1px;transition:background .2s}
.tip-icon:hover{background:#4f46e5}
.cm-tooltip { position:absolute; transform:translate(-50%, -100%); background:rgba(15,23,42,0.95); color:#fff; padding:8px 10px; border-radius:6px; font-size:11px; font-weight:400; line-height:1.45; white-space:normal; width:max-content; max-width:220px; text-transform:none; letter-spacing:normal; text-align:left; pointer-events:none; opacity:0; transition:opacity 0.2s, transform 0.2s; z-index:99999; box-shadow:0 6px 16px rgba(0,0,0,0.2); margin-top:4px; }
.cm-tooltip::after { content:''; position:absolute; top:100%; left:50%; margin-left:-4px; border:4px solid transparent; border-top-color:rgba(15,23,42,0.95); }
.cm-tooltip.show { opacity:1; pointer-events:auto; margin-top:0; }
.cm-tooltip-close { display:inline-block; vertical-align:middle; background:rgba(255,255,255,0.1); border-radius:4px; margin-left:8px; margin-top:-2px; width:16px; height:16px; line-height:14px; text-align:center; cursor:pointer; font-weight:bold; color:#cbd5e1; float:right; }
.cm-tooltip-close:hover { background:rgba(255,255,255,0.25); color:#fff; }

.data-table .td-hint{font-weight:400;font-size:10px;color:var(--faint)}
.data-table .rank-num{text-align:center;font-weight:800;color:var(--indigo);font-size:11px}

.insights-section .insight-group{margin-bottom:12px}
.insight-group-title{display:flex;align-items:center;gap:8px;font-size:12px;font-weight:700;margin-bottom:8px;color:var(--muted)}
.ig-ico{flex-shrink:0;opacity:.85}
.insight-count{margin-left:auto;font-size:10px;background:var(--surface-2);padding:2px 8px;border-radius:999px;color:var(--faint);font-weight:600}
.group-risk .insight-group-title{color:var(--coral)}.group-highlight .insight-group-title{color:var(--teal)}.group-suggestion .insight-group-title{color:var(--indigo)}
.insight-grid{display:grid;gap:6px}
.insight-fold{border:1px solid var(--border);border-radius:var(--radius-sm);background:var(--bg-paper);overflow:hidden;border-left-width:3px;border-left-color:var(--border)}
.insight-tone-danger{border-left-color:var(--coral);background:linear-gradient(90deg,var(--coral-dim),var(--bg-paper))}
.insight-tone-warning{border-left-color:#d97706;background:linear-gradient(90deg,#fffbeb,var(--bg-paper))}
.insight-tone-positive{border-left-color:var(--teal);background:linear-gradient(90deg,var(--teal-dim),var(--bg-paper))}
.insight-tone-info{border-left-color:var(--indigo);background:linear-gradient(90deg,var(--indigo-dim),var(--bg-paper))}
.insight-fold-sum{list-style:none;cursor:pointer;display:flex;align-items:center;justify-content:space-between;gap:10px;padding:8px 10px;font-size:12px;font-weight:700}
.insight-fold-sum::-webkit-details-marker{display:none}
.insight-fold-sum .insight-metric{font-size:13px;font-variant-numeric:tabular-nums;color:var(--ink)}
.insight-fold-body{padding:0 10px 10px;border-top:1px dashed var(--border)}
.insight-body{font-size:12px;color:var(--muted);line-height:1.55;margin-top:8px}
.insight-bench{font-size:10px;color:var(--faint);margin-top:6px;display:block}

.formula-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px}
.formula-grid-dense{gap:6px}
.formula-item{padding:8px 10px;background:var(--surface);border-radius:var(--radius-sm);border:1px solid var(--border);font-size:12px;line-height:1.5}
.formula-item strong{display:block;color:var(--ink);font-size:11px;margin-bottom:2px;font-family:var(--font-serif)}
.formula-item span{color:var(--muted);font-size:11px}

.report-footer{text-align:center;padding:20px 0 12px;font-size:11px;color:var(--faint)}
.footer-line{width:48px;height:2px;background:linear-gradient(90deg,var(--teal),var(--indigo));border-radius:2px;margin:0 auto 12px;opacity:.7}
.footer-note{margin-top:2px;font-size:10px}
.no-data{text-align:center;color:var(--faint);padding:24px;font-size:13px}

@media print{
  body{background:#fff!important;background-image:none!important}
  .report{padding:0;max-width:100%}
  .btn-action{display:none}
  .section,.kpi-card,.report-header{box-shadow:none;break-inside:avoid}
  .report-header{-webkit-print-color-adjust:exact;print-color-adjust:exact}
  details.report-fold,details.insight-fold{display:block!important}
  details.report-fold>summary,details.insight-fold>summary{display:none!important}
  details.report-fold .fold-body,details.insight-fold .insight-fold-body{display:block!important;padding:8px 0}
  .insight-fold{border:none}
}
@media(max-width:900px){
  .kpi-grid{grid-template-columns:repeat(2,1fr)}
  .grid-2,.efficiency-grid,.efficiency-grid-dense,.ranking-grid,.formula-grid{grid-template-columns:1fr}
  .section-strip{text-align:left;margin-left:0;width:100%}
  .report-header h1{font-size:1.25rem}
  .header-actions{position:static;margin-top:10px}
}
@media(max-width:520px){
  .kpi-grid{grid-template-columns:1fr}
  .report{padding:12px}
}</style>`; }

function chartScripts(analysis) {
  const meta = analysis.meta || {};
  const month = meta.month || 1;
  const metrics = getActiveMonths(analysis.workload?.trend || {});
  
  const revData = Array.from({length:12}, (_, i) => analysis.revenue?.trend?.[i+1] || 0);
  const expData = Array.from({length:12}, (_, i) => analysis.expenses?.trend?.[i+1] || 0);
  
  const ef = analysis.efficiency || {};
  const costRatioPct = ef.cost_revenue_ratio != null ? Math.min((ef.cost_revenue_ratio*100).toFixed(1), 100) : 0;
  
  const comp = analysis.expenses?.composition || {};
  const expGroups = [
    {name:"各类物料消耗\n(试剂/耗材/办公)", value: comp["各类消耗"] || 0, itemStyle:{color:"#2563eb"}},
    {name:"人员各项薪酬", value: comp["人员成本"] || 0, itemStyle:{color:"#059669"}},
    {name:"设备折旧与维修", value: comp["设备支出"] || 0, itemStyle:{color:"#d97706"}},
    {name:"第三方委托费\n(外送检验等)", value: comp["第三方费用"] || 0, itemStyle:{color:"#7c3aed"}}
  ];
  const expensePieData = expGroups.filter(g => g.value > 0);
  if (expensePieData.length === 0) expensePieData.push({name:"暂无分类", value:1, itemStyle:{color:"#ccc"}});
  
  const wlData = Array.from({length:12}, (_, i) => {
    let v = analysis.workload?.trend?.[i+1] || 0;
    return {value: v, itemStyle: {color: (i+1)===month ? '#2563eb' : '#93c5fd', borderRadius: [4,4,0,0]}};
  });
  
  const procs = analysis.procurement?.top_items || [];
  const top10 = [...procs].sort((a,b)=>b.amount-a.amount).slice(0,10);
  const procYData = top10.map(p => p.name);
  const procSData = top10.map((p,i) => {
    const colors = ['#2563eb','#059669','#d97706','#7c3aed','#0891b2','#be123c','#4f46e5','#0d9488','#2563eb','#059669'];
    return {value: p.amount, itemStyle: {color: colors[i], borderRadius: [0,4,4,0]}};
  });
  
  const disps = analysis.dispatch?.top_items || [];
  const dispTop10 = [...disps].sort((a,b)=>b.amount-a.amount).slice(0,10);
  const dispYData = dispTop10.map(p => p.name);
  const dispSData = dispTop10.map((p,i) => {
    const colors = ['#dc2626','#ea580c','#d97706','#ca8a04','#65a30d','#16a34a','#0891b2','#0284c7','#4f46e5','#9333ea'];
    return {value: p.amount, itemStyle: {color: colors[i], borderRadius: [0,4,4,0]}};
  });

  return `
<script src="https://cdn.staticfile.net/echarts/5.5.0/echarts.min.js"></script>
<script>
setTimeout(() => {
if (typeof echarts === 'undefined') return console.error('ECharts missed!');
var fmtV=function(v){return v>=10000?(v/10000).toFixed(1)+'万':v>=1000?v.toLocaleString():v};

echarts.init(document.getElementById('chart-revenue-expense')).setOption({
  tooltip:{trigger:'axis',backgroundColor:'rgba(15,23,42,0.92)',borderColor:'transparent',textStyle:{color:'#fff',fontSize:13},
    formatter:function(p){var h=p[0].axisValue;p.forEach(function(i){h+='<br/>'+i.marker+i.seriesName+': '+(i.value>=10000?(i.value/10000).toFixed(2)+'万':i.value.toLocaleString()+'元')});return h}},
  legend:{bottom:0,itemGap:24,textStyle:{fontSize:12,color:'#475569'}},
  grid:{top:20,right:20,bottom:50,left:60,containLabel:false},
  xAxis:{type:'category',data:["1月","2月","3月","4月","5月","6月","7月","8月","9月","10月","11月","12月"],axisLine:{lineStyle:{color:'#e2e8f0'}},axisTick:{show:false},axisLabel:{color:'#64748b'}},
  yAxis:{type:'value',scale:true,axisLabel:{color:'#64748b',formatter:fmtV},splitLine:{lineStyle:{color:'#f1f5f9'}},axisLine:{show:false},axisTick:{show:false}},
  series:[
    {name:'运营收入',type:'line',data:[${revData.join(',')}],
     smooth:true,symbol:'circle',symbolSize:8,lineStyle:{width:2.5,color:'#059669'},
     itemStyle:{color:'#059669'},areaStyle:{color:{type:'linear',x:0,y:0,x2:0,y2:1,colorStops:[{offset:0,color:'rgba(5,150,105,0.12)'},{offset:1,color:'rgba(5,150,105,0)'}]}}},
    {name:'科室支出',type:'line',data:[${expData.join(',')}],
     smooth:true,symbol:'circle',symbolSize:8,lineStyle:{width:2.5,color:'#dc2626'},
     itemStyle:{color:'#dc2626'},areaStyle:{color:{type:'linear',x:0,y:0,x2:0,y2:1,colorStops:[{offset:0,color:'rgba(220,38,38,0.08)'},{offset:1,color:'rgba(220,38,38,0)'}]}}}
  ]
});

echarts.init(document.getElementById('chart-expense-composition')).setOption({
  tooltip:{trigger:'item',formatter:function(p){return p.marker+p.name+'<br/>金额: '+(p.value>=10000?(p.value/10000).toFixed(2)+'万':p.value.toLocaleString()+'元')+'<br/>占比: '+p.percent.toFixed(1)+'%'}},
  legend:{bottom:0,itemGap:16,textStyle:{fontSize:12,color:'#475569'}},
  series:[{
    type:'pie',radius:['38%','62%'],center:['50%','45%'],
    avoidLabelOverlap:true,
    label:{show:true,formatter:'{b}\\n{d}%',fontSize:11,color:'#334155',lineHeight:16},
    labelLine:{show:true,length:12,length2:16,lineStyle:{color:'#94a3b8'}},
    emphasis:{label:{fontSize:13,fontWeight:'bold'},itemStyle:{shadowBlur:10,shadowColor:'rgba(0,0,0,0.15)'}},
    data:${JSON.stringify(expensePieData)}
  }]
});

echarts.init(document.getElementById('chart-workload')).setOption({
  tooltip:{trigger:'axis',formatter:function(p){return p[0].axisValue+'<br/>总门诊: <b>'+p[0].value+'</b> 人次'}},
  grid:{top:20,right:20,bottom:40,left:50},
  xAxis:{type:'category',data:["1月","2月","3月","4月","5月","6月","7月","8月","9月","10月","11月","12月"],axisLine:{lineStyle:{color:'#e2e8f0'}},axisTick:{show:false},axisLabel:{color:'#64748b'}},
  yAxis:{type:'value',axisLabel:{color:'#64748b'},splitLine:{lineStyle:{color:'#f1f5f9'}},axisLine:{show:false},axisTick:{show:false}},
  series:[{type:'bar',barWidth:'50%',data:${JSON.stringify(wlData)},
    label:{show:true,position:'top',fontSize:11,color:'#475569',formatter:'{c}'}}]
});



echarts.init(document.getElementById('chart-procurement-top')).setOption({
  tooltip:{trigger:'axis',axisPointer:{type:'shadow'},
    formatter:function(p){var d=p[0];return d.name+'<br/>金额: <b>'+(d.value>=10000?(d.value/10000).toFixed(2)+'万':d.value.toLocaleString()+'元')+'</b>'}},
  grid:{top:10,right:30,bottom:10,left:10,containLabel:true},
  xAxis:{type:'value',axisLabel:{color:'#64748b',formatter:fmtV},splitLine:{lineStyle:{color:'#f1f5f9'}},axisLine:{show:false},axisTick:{show:false}},
  yAxis:{type:'category',data:${JSON.stringify(procYData)},inverse:true,
    axisLabel:{color:'#334155',fontSize:12,width:120,overflow:'truncate',ellipsis:'..'},
    axisTick:{show:false},axisLine:{lineStyle:{color:'#e2e8f0'}},
    triggerEvent:true},
  series:[{type:'bar',barWidth:'60%',data:${JSON.stringify(procSData)},
    label:{show:true,position:'right',fontSize:11,color:'#475569',formatter:function(p){return p.value>=10000?(p.value/10000).toFixed(1)+'万':p.value.toLocaleString()}}}]
});

echarts.init(document.getElementById('chart-dispatch-top')).setOption({
  tooltip:{trigger:'axis',axisPointer:{type:'shadow'},
    formatter:function(p){var d=p[0];return d.name+'<br/>金额: <b>'+(d.value>=10000?(d.value/10000).toFixed(2)+'万':d.value.toLocaleString()+'元')+'</b>'}},
  grid:{top:10,right:30,bottom:10,left:10,containLabel:true},
  xAxis:{type:'value',axisLabel:{color:'#64748b',formatter:fmtV},splitLine:{lineStyle:{color:'#f1f5f9'}},axisLine:{show:false},axisTick:{show:false}},
  yAxis:{type:'category',data:${JSON.stringify(dispYData)},inverse:true,
    axisLabel:{color:'#334155',fontSize:12,width:120,overflow:'truncate',ellipsis:'..'},
    axisTick:{show:false},axisLine:{lineStyle:{color:'#e2e8f0'}},
    triggerEvent:true},
  series:[{type:'bar',barWidth:'60%',data:${JSON.stringify(dispSData)},
    label:{show:true,position:'right',fontSize:11,color:'#475569',formatter:function(p){return p.value>=10000?(p.value/10000).toFixed(1)+'万':p.value.toLocaleString()}}}]
});

(function(){
  function resizeAll(){ document.querySelectorAll('[id^="chart-"]').forEach(function(el){ var c=echarts.getInstanceByDom(el); if(c) c.resize(); }); }
  window.addEventListener('resize', resizeAll);
  document.querySelectorAll('details.report-fold').forEach(function(d){
    d.addEventListener('toggle', function(){ if(d.open) setTimeout(resizeAll, 80); });
  });

  // Custom tooltips logic
  document.querySelectorAll('.tip-icon, .val-trunc').forEach(function(el) {
    if(!el.getAttribute('title')) return;
    var tip = document.createElement('div');
    tip.className = 'cm-tooltip';
    tip.innerHTML = '<span style="display:inline-block;max-width:calc(100% - 24px);">' + el.getAttribute('title') + '</span><div class="cm-tooltip-close">×</div>';
    el.removeAttribute('title');
    document.body.appendChild(tip);
    
    var hideTip = function() { tip.classList.remove('show'); };
    var showTip = function() {
      document.querySelectorAll('.cm-tooltip.show').forEach(function(t){ if(t!==tip) t.classList.remove('show'); });
      var rect = el.getBoundingClientRect();
      tip.style.left = (rect.left + rect.width/2 + window.scrollX) + 'px';
      tip.style.top = (rect.top + window.scrollY - 6) + 'px';
      tip.classList.add('show');
    };
    
    el.addEventListener('mouseenter', showTip);
    el.addEventListener('mouseleave', function() { setTimeout(function(){ if(!tip.matches(':hover')) hideTip(); }, 150); });
    tip.addEventListener('mouseleave', hideTip);
    el.addEventListener('click', function(e) { e.preventDefault(); e.stopPropagation(); if(tip.classList.contains('show')) hideTip(); else showTip(); });
    tip.querySelector('.cm-tooltip-close').addEventListener('click', function(e) { e.stopPropagation(); hideTip(); });
  });
  document.addEventListener('click', function(e) {
    if(!e.target.closest('.cm-tooltip') && !e.target.closest('.tip-icon') && !e.target.closest('.val-trunc')) {
      document.querySelectorAll('.cm-tooltip.show').forEach(function(t){ t.classList.remove('show'); });
    }
  });

})();

}, 200); // end timeout
</script>`;
}

function reportHeader(analysis) {
  const d = new Date();
  const time = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')} ${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`;
  return `
  <header class="report-header">
    <div class="header-bg"></div>
    <div class="header-content">
      <div class="header-badge">${analysis.meta.year}</div>
      <h1>病理科 运营分析月报</h1>
      <div class="header-meta">
        <span class="header-tag">${analysis.meta.year}年${analysis.meta.month}月</span>
        <span class="header-divider"></span>
        <span class="header-gen">生成于 ${time}</span>
      </div>
    </div>
    <div class="header-actions">
      <button onclick="window.print()" class="btn-action">
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M6 9V2h12v7M6 18H4a2 2 0 01-2-2v-5a2 2 0 012-2h16a2 2 0 012 2v5a2 2 0 01-2 2h-2"></path><rect x="6" y="14" width="12" height="8"></rect></svg>
        打印
      </button>
    </div>
  </header>`;
}

function executiveSummary(analysis) {
  const s = analysis.summary;
  const ef = analysis.efficiency;
  const moMom = analysis.workload?.mom_change;
  
  const costPct = ef.cost_revenue_ratio != null ? (ef.cost_revenue_ratio * 100).toFixed(1) : '0.0';
  const wlQty = analysis.workload?.trend?.[analysis.meta.month] || 0;
  
  const comp = analysis.expenses?.composition || {};
  const incomeStr = fmtMoney(s.total_income);
  const salaryStr = fmtMoney(comp['人员成本'] || 0);
  const surplusStr = fmtMoney(s.operation_surplus);
  
  const consumableStr = fmtMoney(comp['各类消耗'] || 0);
  const procAmount = analysis.procurement?.total_amount || 0;
  
  let aiText = `主任您好：经系统核算，本周期科室各项<strong>门诊运营总收入共 ${fmtMoney(s.total_income)}</strong>，归集的<strong>总考核支出共 ${fmtMoney(s.total_expense)}</strong>。在细分支出（其中人员薪酬发放 <strong>${salaryStr}</strong>，其余为设备折旧与日常消耗）扣除后，科室本月最终实现利润结余 <strong>${surplusStr}</strong>。`;
  
  if ((s.operation_surplus||0) < 0) {
    aiText += ' <span style="color:var(--coral);font-weight:600">由于目前处于亏损状态，需提请重点剖析亏损源头（如耗材占比失调或收费结构问题）。</span>';
  } else {
    aiText += ' <span style="color:var(--teal);font-weight:600">整体财务健康度保持稳健。</span>';
  }
  
  aiText += `<br><span style="display:inline-block;margin-top:8px">物资采购动态：本月各类消耗领用账面扣除金额为 <strong>${consumableStr}</strong>，而科室实际新发起的采购打单金额为 <strong>${fmtMoney(procAmount)}</strong>。</span>`;
  
  const cVal = comp['各类消耗'] || 0;
  if (cVal > 0) {
    if (procAmount > cVal * 2) {
      aiText += `<span style="color:var(--indigo);font-weight:600"> 实际采购大幅高于当月消耗。这通常意味着科室本月进行了【季度/周期性大宗备库】。请同步关注大额入库试剂的效期风险。</span>`;
    } else if (procAmount < cVal * 0.3) {
      aiText += `<span style="color:#d97706;font-weight:600"> 采购额远低于本月账面消耗。若科室采取【季度集中采购】模式，此属于正常资金流波谷；否则请务必提醒护士长或库管实地清点核心检测试剂，视情况及时打单补货，防范断货风险。</span>`;
    } else {
      aiText += `<span style="color:var(--muted)"> 资金流出总量与耗材消耗基准基本吻合，库存周转处于良性动态平衡。</span>`;
    }
  }

  return `
  <section class="exec-summary" aria-label="运营概览">
    <div class="exec-summary-head">
      <h2 class="exec-title">智能财物简报</h2>
      <span class="exec-sub">Executive Summary</span>
    </div>
    
    <div class="ai-summary">
      <div class="ai-summary-icon">
        <svg fill="none" viewBox="0 0 24 24" stroke-width="1.5" stroke="currentColor" style="width:24px;height:24px;">
          <path stroke-linecap="round" stroke-linejoin="round" d="M9.813 15.904 9 18.75l-.813-2.846a4.5 4.5 0 0 0-3.09-3.09L2.25 12l2.846-.813a4.5 4.5 0 0 0 3.09-3.09L9 5.25l.813 2.846a4.5 4.5 0 0 0 3.09 3.09L15.75 12l-2.846.813a4.5 4.5 0 0 0-3.09 3.09ZM18.259 8.715 18 9.75l-.259-1.035a3.375 3.375 0 0 0-2.455-2.456L14.25 6l1.036-.259a3.375 3.375 0 0 0 2.455-2.456L18 2.25l.259 1.035a3.375 3.375 0 0 0 2.456 2.456L21.75 6l-1.035.259a3.375 3.375 0 0 0-2.456 2.456ZM16.894 20.567 16.5 21.75l-.394-1.183a2.25 2.25 0 0 0-1.423-1.423L13.5 18.75l1.183-.394a2.25 2.25 0 0 0 1.423-1.423l.394-1.183.394 1.183a2.25 2.25 0 0 0 1.423 1.423l1.183.394-1.183.394a2.25 2.25 0 0 0-1.423 1.423Z" />
        </svg>
      </div>
      <div class="ai-summary-text">${aiText}</div>
    </div>
    
    <div class="kpi-grid">
      <div class="kpi-card">
        <div class="kpi-icon kpi-icon-income"><svg viewBox="0 0 24 24"><path fill="currentColor" d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1.41 16.09V20h-2.67v-1.93c-1.71-.36-3.16-1.46-3.27-3.4h1.96c.1 1.05.82 1.87 2.65 1.87 1.96 0 2.4-.98 2.4-1.59 0-.83-.44-1.61-2.67-2.14-2.48-.6-4.18-1.62-4.18-3.67 0-1.72 1.39-2.84 3.11-3.21V4h2.67v1.95c1.86.45 2.79 1.86 2.85 3.39H14.3c-.05-1.11-.64-1.87-2.22-1.87-1.5 0-2.4.68-2.4 1.64 0 .84.65 1.39 2.67 1.91s4.18 1.39 4.18 3.91c-.01 1.83-1.38 2.83-3.12 3.16z"></path></svg></div>
        <div class="kpi-content">
          <div class="kpi-label">运营收入</div>
          <div class="kpi-value">${fmtMoney(s.total_income)}</div>
        </div>
      </div>
      <div class="kpi-card">
        <div class="kpi-icon kpi-icon-expense"><svg viewBox="0 0 24 24"><path fill="currentColor" d="M11.8 10.9c-2.27-.59-3-1.2-3-2.15 0-1.09 1.01-1.85 2.7-1.85 1.78 0 2.44.85 2.5 2.1h2.21c-.07-1.72-1.12-3.3-3.21-3.81V3h-3v2.16c-1.94.42-3.5 1.68-3.5 3.61 0 2.31 1.91 3.46 4.7 4.13 2.5.6 3 1.48 3 2.41 0 .69-.49 1.79-2.7 1.79-2.06 0-2.87-.92-2.98-2.1h-2.2c.12 2.19 1.76 3.42 3.68 3.83V21h3v-2.15c1.95-.37 3.5-1.5 3.5-3.55 0-2.84-2.43-3.81-4.7-4.4z"></path></svg></div>
        <div class="kpi-content">
          <div class="kpi-label">科室支出</div>
          <div class="kpi-value">${fmtMoney(s.total_expense)}</div>
        </div>
      </div>
      <div class="kpi-card">
        <div class="kpi-icon kpi-icon-surplus"><svg viewBox="0 0 24 24"><path fill="currentColor" d="M16 6l2.29 2.29-4.88 4.88-4-4L2 16.59 3.41 18l6-6 4 4 6.3-6.29L22 12V6h-6z"></path></svg></div>
        <div class="kpi-content">
          <div class="kpi-label">运营结余</div>
          <div class="kpi-value ${changeClass(s.operation_surplus)}">${fmtMoney(s.operation_surplus)}</div>
        </div>
      </div>
      <div class="kpi-card kpi-gauge">
        <div class="kpi-icon kpi-icon-ratio"><svg viewBox="0 0 24 24"><path fill="currentColor" d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-7 3c1.93 0 3.5 1.57 3.5 3.5S13.93 13 12 13s-3.5-1.57-3.5-3.5S10.07 6 12 6zm7 13H5v-.23c0-.62.36-1.2.97-1.63C7.25 15.8 9.47 15 12 15s4.75.8 6.03 2.14c.59.43.97 1.01.97 1.63V19z"></path></svg></div>
        <div class="kpi-content">
          <div class="kpi-label">支收比<span class="tip-icon" title="科室全成本总支出 ÷ 门诊与病房计奖收入总和。反映科室每赚1元钱需要花多少钱，基准安全线≤70%">?</span></div>
          <div class="kpi-value">${costPct}%</div>
          <div class="kpi-sub ${costPct<=70?'':'val-neg'}">≤ 基准70% ${costPct<=70?'✓':'✘'}</div>
        </div>
        <div class="gauge-bar"><div class="gauge-fill" style="width:${Math.min(costPct,100)}%;background:#059669"></div><div class="gauge-marks"><span style="left:70%"></span><span style="left:80%" class="mark-danger"></span></div></div>
      </div>
      <div class="kpi-card">
        <div class="kpi-icon kpi-icon-workload"><svg viewBox="0 0 24 24"><path fill="currentColor" d="M16 11c1.66 0 2.99-1.34 2.99-3S17.66 5 16 5s-3 1.34-3 3 1.34 3 3 3zm-8 0c1.66 0 2.99-1.34 2.99-3S9.66 5 8 5 5 6.34 5 8s1.34 3 3 3zm0 2c-2.33 0-7 1.17-7 3.5V19h14v-2.5c0-2.33-4.67-3.5-7-3.5zm8 0c-.29 0-.62.02-.97.05 1.16.84 1.97 1.97 1.97 3.45V19h6v-2.5c0-2.33-4.67-3.5-7-3.5z"></path></svg></div>
        <div class="kpi-content">
          <div class="kpi-label">总门诊量</div>
          <div class="kpi-value">${wlQty} <small>人次</small></div>
          <div class="kpi-sub ${changeClass(moMom, true)}">环比 ${moMom>0?'↑':'↓'} ${fmtChange(moMom)}</div>
        </div>
      </div>
      <div class="kpi-card">
        <div class="kpi-icon kpi-icon-percapita"><svg viewBox="0 0 24 24"><path fill="currentColor" d="M12 12c2.21 0 4-1.79 4-4s-1.79-4-4-4-4 1.79-4 4 1.79 4 4 4zm0 2c-2.67 0-8 1.34-8 4v2h16v-2c0-2.66-5.33-4-8-4z"></path></svg></div>
        <div class="kpi-content">
          <div class="kpi-label">人均产出</div>
          <div class="kpi-value">${fmtMoney(ef.per_capita_income)}</div>
          <div class="kpi-sub">人均月收入</div>
        </div>
      </div>
      <div class="kpi-card">
        <div class="kpi-icon kpi-icon-satisfaction"><svg viewBox="0 0 24 24"><path fill="currentColor" d="M11.99 2C6.47 2 2 6.48 2 12s4.47 10 9.99 10C17.52 22 22 17.52 22 12S17.52 2 11.99 2zM12 20c-4.42 0-8-3.58-8-8s3.58-8 8-8 8 3.58 8 8-3.58 8-8 8zm3.5-9c.83 0 1.5-.67 1.5-1.5S16.33 8 15.5 8 14 8.67 14 9.5s.67 1.5 1.5 1.5zm-7 0c.83 0 1.5-.67 1.5-1.5S9.33 8 8.5 8 7 8.67 7 9.5 7.67 11 8.5 11zm3.5 6.5c2.33 0 4.31-1.46 5.11-3.5H6.89c.8 2.04 2.78 3.5 5.11 3.5z"></path></svg></div>
        <div class="kpi-content">
          <div class="kpi-label">满意度</div>
          <div class="kpi-value">${s.satisfaction_rank ? '#' + s.satisfaction_rank : '--'}</div>
          <div class="kpi-sub">全院数据排名</div>
        </div>
      </div>
      <div class="kpi-card">
        <div class="kpi-icon kpi-icon-team"><svg viewBox="0 0 24 24"><path fill="currentColor" d="M16 11c1.66 0 2.99-1.34 2.99-3S17.66 5 16 5c-1.66 0-3 1.34-3 3s1.34 3 3 3zm-8 0c1.66 0 2.99-1.34 2.99-3S9.66 5 8 5C6.34 5 5 6.34 5 8s1.34 3 3 3zm0 2c-2.33 0-7 1.17-7 3.5V19h14v-2.5c0-2.33-4.67-3.5-7-3.5zm8 0c-.29 0-.62.02-.97.05 1.16.84 1.97 1.97 1.97 3.45V19h6v-2.5c0-2.33-4.67-3.5-7-3.5z"></path></svg></div>
        <div class="kpi-content">
          <div class="kpi-label">团队人数</div>
          <div class="kpi-value">${fmtNum(analysis.team?.count || 47)} <small>人</small></div>
        </div>
      </div>
    </div>
  </section>`;
}

function renderIndicatorTable(items, months, showChange, cm) {
  if (!items?.length) return '<p class="no-data">暂无数据</p>';
  return `
  <div class="table-scroller">
    <table class="data-table">
      <thead>
        <tr>
          <th class="st-col st-1 c">指标</th><th class="c st-col st-2">上年均值</th><th class="c st-col st-3 st-divider">本年均值</th>
          ${months.map((m, idx) => `<th class="c ${(idx+1)===cm ? 'col-highlight' : ''}">${m}月</th>`).join('')}
          ${showChange ? '<th class="c st-col st-r1">环比<span class="tip-icon" title="本月该项数据与【上月各项数据】直接对比的涨跌幅">?</span></th><th class="c st-col st-r2">同比<span class="tip-icon" title="本月该项数据与【去年均值】相比的涨跌幅">?</span></th>' : ''}
        </tr>
      </thead>
      <tbody>
        ${items.map(i => `
        <tr>
          <td class="td-name st-col st-1">${esc(i.label)}${(function(l){
            if(l.includes('医事服务费')) return '<span class="tip-icon" title="按本科室执行的核算提取规则：检查化验类30%，手术麻醉类50%，其他类100%，且剔除药耗。">?</span>';
            if(l.includes('运营收入合计')) return '<span class="tip-icon" title="科室全口径实际业务收入：包含计奖收入 + 折算后的其他收入分摊比例 + 医事服务费。">?</span>';
            if(l.includes('计奖收入')) return '<span class="tip-icon" title="纯医疗项目服务类开单与执行产生的直接绩效基数收入，不含药耗材料。">?</span>';
            return '';
          })(i.label)}</td>
          <td class="c num st-col st-2">${i.prev_year_avg == null ? '0' : fmtNum(i.prev_year_avg)}</td>
          <td class="c num st-col st-3 st-divider">${i.current_year_avg == null ? '0' : fmtNum(i.current_year_avg)}</td>
          ${months.map((m, idx) => `<td class="c num ${(idx+1)===cm ? 'col-highlight' : ''}">${i.monthly[m] == null ? '0' : fmtNum(i.monthly[m])}</td>`).join('')}
          ${showChange ? `<td class="c st-col st-r1 ${changeClass(i.mom_change)}">${fmtChange(i.mom_change)}</td><td class="c st-col st-r2 ${changeClass(i.yoy_change)}">${fmtChange(i.yoy_change)}</td>` : ''}
        </tr>`).join('')}
      </tbody>
    </table>
  </div>`;
}

function financialSection(analysis) {
  const { revenue, expenses, summary } = analysis;
  const mo = analysis.meta.month;
  
  return `
  <section class="section section-compact">
    <div class="section-header section-header-inline">
      <h2>一、经济运行</h2>
      <span class="section-desc">Financial</span>
      <div class="section-strip">
        <span class="strip-item">运营收入<span class="tip-icon" title="门诊计奖收入与门诊医事服务费等各项收入之和。">?</span> <strong>${fmtMoney(summary.total_income)}</strong></span><span class="strip-sep">·</span>
        <span class="strip-item">支出 <strong>${fmtMoney(summary.total_expense)}</strong></span><span class="strip-sep">·</span>
        <span class="strip-item">结余 <strong class="${changeClass(summary.operation_surplus)}">${fmtMoney(summary.operation_surplus)}</strong></span>
      </div>
    </div>
    <div class="grid-2 grid-tight">
      <div class="chart-card chart-card-sm">
        <h3>收入支出趋势</h3>
        <div id="chart-revenue-expense" class="chart-box" style="height: 220px;"></div>
      </div>
      <div class="chart-card chart-card-sm">
        <h3>支出结构 ${mo}月</h3>
        <div id="chart-expense-composition" class="chart-box" style="height: 260px;"></div>
      </div>
    </div>
    <div class="mini-cards mini-cards-row">
      <div class="mini-card sm"><span class="mini-label">收入同比<span class="tip-icon" title="【同比】指本月运营收入与「去年均值」相比的涨跌情况。">?</span></span><span class="mini-value ${changeClass(summary.income_yoy_change)}">${fmtChange(summary.income_yoy_change)}</span></div>
      <div class="mini-card sm"><span class="mini-label">支出同比</span><span class="mini-value ${changeClass(summary.expense_yoy_change, true)}">${fmtChange(summary.expense_yoy_change)}</span></div>
      <div class="mini-card sm"><span class="mini-label">运营结余</span><span class="mini-value ${changeClass(summary.operation_surplus)}">${fmtMoney(summary.operation_surplus)}</span></div>
    </div>
    <details class="report-fold">
      <summary class="fold-summary"><span class="fold-chevron"></span><span class="fold-title-text">逐项明细表（收入 · 支出）</span><span class="fold-hint">展开月结表</span></summary>
      <div class="fold-body">
        <div class="table-wrap tight"><h4 class="table-inline-title">收入明细</h4>
        ${renderIndicatorTable(revenue.items, ML.map((_, i) => i + 1), true, mo)}
        </div>
        <div class="table-wrap tight"><h4 class="table-inline-title">支出明细</h4>
        ${renderIndicatorTable(expenses.items, ML.map((_, i) => i + 1), true, mo)}
        </div>
      </div>
    </details>
  </section>`;
}

function efficiencySection(analysis) {
  const ef = analysis.efficiency;
  const costRatioPct = ef.cost_revenue_ratio != null ? (ef.cost_revenue_ratio * 100).toFixed(1) + '%' : '--';
  return `
  <section class="section section-compact">
    <div class="section-header section-header-inline">
      <h2>二、运营效率</h2>
      <span class="section-desc">Efficiency</span>
      <div class="section-strip"><span class="strip-item">支收比 <strong>${costRatioPct}</strong></span></div>
    </div>
    <div class="efficiency-grid">
      <div class="eff-card eff-${statusOf(ef.cost_revenue_ratio, 0.7, 0.8)}">
        <div class="eff-status"></div><div class="eff-label">支收比<span class="tip-icon" title="科室全成本总支出 ÷ 运营收入总和。反映科室每赚1元钱需要花多少钱，基准安全线≤70%">?</span></div>
        <div class="eff-value">${costRatioPct}</div><div class="eff-bench">≤70%</div>
      </div>
      <div class="eff-card eff-${statusOf(ef.personnel_cost_ratio, 0.35, 0.40)}">
        <div class="eff-status"></div><div class="eff-label">人力成本占比<span class="tip-icon" title="人员工资及绩效发放 ÷ 运营总收入。体现科室劳动创造比率。">?</span></div>
        <div class="eff-value">${fmtPct(ef.personnel_cost_ratio)}</div><div class="eff-bench">≤35%</div>
      </div>
      <div class="eff-card eff-${statusOf(ef.consumable_per_100_income, 25, 30)}">
        <div class="eff-status"></div><div class="eff-label">百元收入卫材<span class="tip-icon" title="当期卫材消耗账面金额 ÷ (运营总收入÷100)。管控目标通常为25-30元以内。">?</span></div>
        <div class="eff-value">${fmtMoney(ef.consumable_per_100_income)}</div><div class="eff-bench">≤25元</div>
      </div>
      <div class="eff-card eff-${statusOf(ef.operation_margin, 0.1, 0, true)}">
        <div class="eff-status"></div><div class="eff-label">运营利润率<span class="tip-icon" title="科室最终运营结余 ÷ 运营总收入。体现科室实际纯利水平。">?</span></div>
        <div class="eff-value">${fmtPct(ef.operation_margin)}</div><div class="eff-bench">&gt;10%</div>
      </div>
      <div class="eff-card eff-${statusOf(ef.award_cost_ratio, 0.7, 0.8)}">
        <div class="eff-status"></div><div class="eff-label">奖金支收比<span class="tip-icon" title="科室支出合计 ÷ 科室计奖收入合计。以计奖收入为基准的支收比，反映奖金分配视角下科室的成本消耗水平，是医院绩效考核的核心指标。">?</span></div>
        <div class="eff-value">${fmtPct(ef.award_cost_ratio)}</div><div class="eff-bench">≤70%</div>
      </div>
      <div class="eff-card eff-${ef.per_capita_surplus > 0 ? 'good' : 'bad'}">
        <div class="eff-status"></div><div class="eff-label">人均月结余<span class="tip-icon" title="最终利润结余 ÷ 科室总在职人数。体现全员平均创造的正收益。">?</span></div>
        <div class="eff-value">${fmtMoney(ef.per_capita_surplus)}</div>
      </div>
    </div>
  </section>`;
}

function workloadSection(analysis) {
  const wl = analysis.workload;
  const mo = analysis.meta.month;
  const currWorkload = wl?.trend?.[mo] || 0;
  return `
  <section class="section section-compact">
    <div class="section-header section-header-inline">
      <h2>三、工作量</h2>
      <span class="section-desc">Workload</span>
      <div class="section-strip">
        <span class="strip-item">合计 <strong>${currWorkload}</strong> 人次</span><span class="strip-sep">·</span>
        <span class="strip-item">环比<span class="tip-icon" title="【环比】指本月数据与「上相邻的一个月」相比较的涨跌幅度（而非去年同月）。">?</span> <strong class="${changeClass(wl.mom_change)}">${fmtChange(wl.mom_change)}</strong></span>
      </div>
    </div>
    <div class="grid-2 grid-tight">
      <div class="chart-card chart-card-sm">
        <h3>门诊趋势</h3>
        <div id="chart-workload" class="chart-box" style="height: 220px;"></div>
      </div>
      <div class="stat-panel stat-panel-dense">
        <h3>本月门诊量构成</h3>
        <div class="stat-row"><span>总计量</span><span class="stat-big">${currWorkload}<small>人次</small></span></div>
        <div class="stat-row"><span>人均负荷</span><span class="stat-big">${((currWorkload / (analysis.team?.count || 47))||0).toFixed(1)}<small>人次/人</small></span></div>
        <div class="stat-divider"></div>
        <div class="stat-row sub"><span>专家门诊量</span><span>0</span></div>
        <div class="stat-row sub"><span>特需门诊量</span><span>0</span></div>
        <div class="stat-row sub"><span>普通门诊量</span><span>${currWorkload}</span></div>
      </div>
    </div>
    <details class="report-fold">
      <summary class="fold-summary"><span class="fold-chevron"></span><span class="fold-title-text">工作量明细</span></summary>
      <div class="fold-body">
        <div class="table-wrap tight"><h4 class="table-inline-title">逐月分布</h4>
        ${renderIndicatorTable(wl.items, ML.map((_, i) => i + 1), false, mo)}
        </div>
      </div>
    </details>
  </section>`;
}

function rankingsSection(analysis) {
  const r = analysis.rankings;
  if (!r) return '';
  return `
  <section class="section section-compact">
    <div class="section-header section-header-inline">
      <h2>四、全院排名</h2>
      <span class="section-desc">Rankings</span>
      <div class="section-strip"><span class="strip-item">增量总排名 <strong>第${r.change_rank||'-'}名</strong></span></div>
    </div>
    <div class="ranking-grid ranking-grid-dense">
      <div class="rank-card">
        <div class="rank-label">患者满意度<span class="tip-icon" title="基于当月出院及门诊的患者医患关系问卷口碑转化评分（满分5分）。">?</span></div>
        <div class="rank-value">${r.satisfaction_score != null ? Number(r.satisfaction_score).toFixed(2) : '--'}</div>
      </div>
      <div class="rank-card">
        <div class="rank-label">满意度排名</div>
        <div class="rank-value">第${r.satisfaction_rank||'-'}名</div>
      </div>
      <div class="rank-card">
        <div class="rank-label">科室综合增量<span class="tip-icon" title="评估科室支收比改善及业务量环比提升的综合考核系数。用以衡量当月管理进步幅度。">?</span></div>
        <div class="rank-value">第${r.change_rank||'-'}名</div>
        <div class="rank-badge">#${r.change_rank||'-'}</div>
      </div>
    </div>
  </section>`;
}

function materialsSection(analysis) {
  const p = analysis.procurement || {};
  const d = analysis.dispatch || {};
  const totalProc = p.total_amount || 0;
  const totalDisp = d.total_amount || 0;
  const suppliers = p.supplier_distribution ? Object.keys(p.supplier_distribution).length : 0;
  
  const procItems = p.all_items || [];
  const procRows = procItems.map((v, i) => {
    const price = v.qty ? fmtMoney(v.amount / v.qty) : '-';
    // Use td-ellipsis for truncation and val-trunc for custom hover tooltip
    const safeName = (v.name||'').replace(/"/g, '&quot;');
    return `<tr><td class="r rank-num">${i+1}</td><td class="td-name"><div class="td-ellipsis val-trunc" style="max-width: 220px; display: inline-block; vertical-align: middle; border-bottom: none;" title="${safeName}">${v.name}</div></td><td class="r num">${v.qty||'-'}</td><td class="r num">${price}</td><td class="r num">${fmtMoney(v.amount)}</td></tr>`;
  }).join('');
  
  const dispItems = d.all_items || [];
  const dispRows = dispItems.map((v, i) => {
    const price = v.qty ? fmtMoney(v.amount / v.qty) : '-';
    const safeName = (v.name||'').replace(/"/g, '&quot;');
    return `<tr><td class="r rank-num">${i+1}</td><td class="td-name"><div class="td-ellipsis val-trunc" style="max-width: 220px; display: inline-block; vertical-align: middle; border-bottom: none;" title="${safeName}">${v.name}</div></td><td class="r num">${v.qty||'-'}</td><td class="r num">${price}</td><td class="r num">${fmtMoney(v.amount)}</td></tr>`;
  }).join('');

  return `
  <section class="section section-compact">
    <div class="section-header section-header-inline">
      <h2>五、出库入库明细</h2>
      <span class="section-desc">Inventory</span>
      <div class="section-strip">
        <span class="strip-item">采 <strong>${fmtMoney(totalProc)}</strong></span><span class="strip-sep">·</span>
        <span class="strip-item">发 <strong>${fmtMoney(totalDisp)}</strong></span><span class="strip-sep">·</span>
        <span class="strip-item"><strong>${suppliers}</strong> 供货渠道</span>
      </div>
    </div>
    <div class="mini-cards mini-cards-row">
      <div class="mini-card sm"><span class="mini-label">当月采购</span><span class="mini-value">${fmtMoney(totalProc)}</span></div>
      <div class="mini-card sm"><span class="mini-label">当月出库</span><span class="mini-value">${fmtMoney(totalDisp)}</span></div>
      <div class="mini-card sm"><span class="mini-label">合作商</span><span class="mini-value">${suppliers}</span><span class="mini-sub">家</span></div>
    </div>
    <div class="grid-2 grid-tight">
      <div class="chart-card chart-card-sm">
        <h3>采购入库 TOP 10</h3>
        <div id="chart-procurement-top" class="chart-box" style="height: 280px;"></div>
      </div>
      <div class="chart-card chart-card-sm">
        <h3>领用出库 TOP 10</h3>
        <div id="chart-dispatch-top" class="chart-box" style="height: 280px;"></div>
      </div>
    </div>

    <details class="report-fold" style="margin-top:20px;">
      <summary class="fold-summary"><span class="fold-chevron"></span><span class="fold-title-text">参阅：当月出库入库完整明细表</span><span class="fold-hint">点击展开明细</span></summary>
      <div class="fold-body">
        <div class="grid-2 grid-tight" style="gap:20px;">
          <div class="table-wrap tight">
            <div class="table-inline-title">当月采购入库账单明细</div>
            <div class="table-scroller">
              <table class="data-table">
                <thead><tr><th style="width:40px" class="c">排名</th><th>品名</th><th class="r">数量</th><th class="r">单价</th><th class="r">总金额</th></tr></thead>
                <tbody>${procRows || '<tr><td colspan="5" class="c td-hint">暂无记录</td></tr>'}</tbody>
              </table>
            </div>
          </div>
          <div class="table-wrap tight">
            <div class="table-inline-title">当月领用出库账单明细</div>
            <div class="table-scroller">
              <table class="data-table">
                <thead><tr><th style="width:40px" class="c">排名</th><th>品名</th><th class="r">数量</th><th class="r">单价</th><th class="r">总金额</th></tr></thead>
                <tbody>${dispRows || '<tr><td colspan="5" class="c td-hint">暂无记录</td></tr>'}</tbody>
              </table>
            </div>
          </div>
        </div>
      </div>
    </details>
  </section>`;
}

function appendixSection() {
  return `
  <section class="section section-compact appendix-wrap">
    <details class="report-fold">
      <summary class="fold-summary"><span class="fold-chevron"></span><span class="fold-title-text">参阅：科室精细化核算规则</span><span class="fold-hint">点击展开全文明细</span></summary>
      <div class="fold-body">
        <div class="formula-grid formula-grid-dense" style="display:flex; flex-direction:column; gap:8px;">
          <div class="formula-item"><strong>运营收入合计</strong><span style="white-space:normal;line-height:1.5;">门诊计奖收入 + 医事服务费（不含药耗，其他收入按30%/70%计入开单科室及执行科室） + 病房计奖收入 + 医事服务费（不含药耗，检查化验收入按30%，手术费及麻醉50%，其他收入100%计入科室）</span></div>
          <div class="formula-item"><strong>运营支出合计</strong><span style="white-space:normal;line-height:1.5;">人员支出 + 设备支出 + 非收费医材支出 + 其他消耗支出 + 第三方服务支出</span></div>
          <div class="formula-item"><strong>其他消耗支出</strong><span style="white-space:normal;line-height:1.5;">试剂 + 设备类耗材 + 办公耗材 + 医疗废弃物 + 氧气</span></div>
          <div class="formula-item"><strong>设备支出合计</strong><span style="white-space:normal;line-height:1.5;">设备折旧 + 设备维修</span></div>
          <div class="formula-item"><strong>设备折旧</strong><span style="white-space:normal;line-height:1.5;">（资金来源为自筹）= 设备上账金额 / 60，按月扣除；六到十年逐年递减</span></div>
          <div class="formula-item"><strong>设备维修</strong><span style="white-space:normal;line-height:1.5;">（资金来源为自筹及财政）= 设备上账金额 * 0.005，按月扣除</span></div>
        </div>
      </div>
    </details>
  </section>`;
}

function generateReportBody(analysis) {
  // CRITICAL: We MUST include the CSS and Scripts dynamically if this body is embedded.
  // We can include a marker variable or just output it directly so the index.ejs uses it instantly.
  return `
  ${reportCSS()}
  <div class="report-container">
    ${reportHeader(analysis)}
    ${executiveSummary(analysis)}
    <div class="report-body">
      ${financialSection(analysis)}
      ${efficiencySection(analysis)}
      ${workloadSection(analysis)}
      ${rankingsSection(analysis)}
      ${materialsSection(analysis)}
      ${appendixSection()}
    </div>
    <div style="height:40px"></div>
  </div>
  ${chartScripts(analysis)}
  `;
}

function generateReport(analysis) {
  return `<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8"><title>运营分析报告</title><style>body{background:#f7f8fa;margin:0}</style></head><body>${generateReportBody(analysis)}</body></html>`;
}

module.exports = { generateReport, generateReportBody, chartScripts, reportCSS };
