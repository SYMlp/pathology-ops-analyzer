function esc(str) {
  return (str || '').replace(/[&<>'"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' })[c]);
}

function fmtInt(n) {
  if (n == null || Number.isNaN(n)) return '—';
  return Math.round(n).toLocaleString('zh-CN');
}

function fmtMoney(n) {
  if (n == null || Number.isNaN(n)) return '—';
  return n.toLocaleString('zh-CN', { maximumFractionDigits: 0 });
}

function workloadCSS() {
  return `<style>
:root{--ink:#0f172a;--muted:#64748b;--border:#e2e8f0;--bg:#f8fafc;--card:#fff;--blue:#2563eb;--teal:#0d9488}
body{margin:0;font-family:'Inter','Noto Sans SC',system-ui,sans-serif;background:var(--bg);color:var(--ink);font-size:14px;line-height:1.5}
.wl-report{max-width:1100px;margin:0 auto;padding:24px 20px 40px}
.wl-h1{font-size:1.35rem;font-weight:800;margin:0 0 6px}
.wl-meta{color:var(--muted);font-size:13px;margin-bottom:28px}
.wl-section{background:var(--card);border:1px solid var(--border);border-radius:12px;padding:18px 20px;margin-bottom:20px;box-shadow:0 1px 3px rgba(0,0,0,.06)}
.wl-section h2{font-size:1.05rem;margin:0 0 14px;font-weight:700;border-bottom:1px solid var(--border);padding-bottom:10px}
.wl-table{width:100%;border-collapse:collapse;font-size:13px}
.wl-table th,.wl-table td{border:1px solid var(--border);padding:8px 10px;text-align:left}
.wl-table th{background:#f1f5f9;font-weight:600}
.wl-table td.num{text-align:right;font-variant-numeric:tabular-nums}
.wl-chart{height:380px;width:100%;min-height:300px;margin-top:24px}
.wl-chart-sm{height:320px}
.wl-section .wl-chart:first-of-type{margin-top:18px}
.wl-actions{display:flex;flex-wrap:wrap;gap:8px;margin:6px 0 12px}
.wl-actions button{font-size:12px;padding:6px 12px;border:1px solid var(--border);border-radius:8px;background:#fff;cursor:pointer}
.wl-actions button:hover{background:#f1f5f9}
.wl-detail-cap{margin:18px 0 6px;font-size:13px;font-weight:600;color:var(--muted)}
@media(max-width:640px){.wl-chart{height:300px}.wl-chart-sm{height:260px}}
</style>`;
}

function workloadTableHTML(data) {
  const w = data.workload;
  if (!w?.items?.length) return '<p class="wl-meta">未解析到工作量数据（请检查「工作量统计」Sheet）。</p>';

  let rows = '';
  for (const it of w.items) {
    rows += `<tr>
      <td>${esc(it.name)}</td>
      <td class="num">${fmtInt(it.cases)}</td>
      <td class="num">${fmtInt(it.blocks)}</td>
      <td class="num">${fmtInt(it.frozen_cases)}</td>
      <td class="num">${fmtInt(it.frozen_blocks)}</td>
    </tr>`;
  }

  return `<table class="wl-table">
    <thead><tr><th>名称</th><th>例数</th><th>块数</th><th>冰冻例数</th><th>冰冻块数</th></tr></thead>
    <tbody>${rows}</tbody>
  </table>`;
}

function revenueTableHTML(data) {
  const r = data.revenue;
  if (!r) return '<p class="wl-meta">未解析到收入报表（「收入金额」Sheet）。</p>';

  let extra = '';
  if (r.outpatientDetails?.length) {
    let dRows = '';
    for (const d of r.outpatientDetails) {
      const qty = d.total?.qty ?? null;
      const ta = d.total?.amount ?? null;
      const tza = d.tongzhou?.amount ?? null;
      const qtyCell = qty ? fmtInt(qty) : '—';
      dRows += `<tr><td>${d.price}元档</td><td class="num">${qtyCell}</td><td class="num">${fmtMoney(ta)}</td><td class="num">${fmtMoney(tza)}</td></tr>`;
    }
    extra = `<div class="wl-detail-cap">门诊具体档位</div><table class="wl-table">
      <thead><tr><th>档位</th><th>数量</th><th>金额</th><th>通州</th></tr></thead><tbody>${dRows}</tbody></table>`;
  }

  return `<table class="wl-table">
    <thead><tr><th>项目</th><th>数量</th><th>金额</th><th>备注（通州）</th></tr></thead>
    <tbody>
      <tr><td>病房</td><td class="num">—</td><td class="num">${fmtMoney(r.inpatient?.total)}</td><td class="num">${fmtMoney(r.inpatient?.tongzhou)}</td></tr>
      <tr><td>门诊</td><td class="num">—</td><td class="num">${fmtMoney(r.outpatient?.total)}</td><td class="num">${fmtMoney(r.outpatient?.tongzhou)}</td></tr>
      <tr><td>病理会诊（门诊）</td><td class="num">—</td><td class="num">${fmtMoney(r.consultOutpatient?.total)}</td><td class="num">${fmtMoney(r.consultOutpatient?.tongzhou)}</td></tr>
      <tr><td>病理会诊（住院）</td><td class="num">—</td><td class="num">${fmtMoney(r.consultInpatient?.total)}</td><td class="num">${fmtMoney(r.consultInpatient?.tongzhou)}</td></tr>
      <tr><td><strong>合计</strong></td><td class="num">—</td><td class="num"><strong>${fmtMoney(r.grandTotal?.total)}</strong></td><td class="num"><strong>${fmtMoney(r.grandTotal?.tongzhou)}</strong></td></tr>
    </tbody>
  </table>${extra}`;
}

function consultationTableHTML(data) {
  const c = data.consultation;
  if (!c?.outpatient) return '<p class="wl-meta">未解析到会诊统计（「会诊统计」Sheet）。</p>';

  const op = c.outpatient;
  const ip = c.inpatient;
  const row = (label, xz, tz) =>
    `<tr><td>${esc(label)}</td><td class="num">${xz}</td><td class="num">${tz}</td><td class="num">—</td></tr>`;

  return `<table class="wl-table">
    <thead><tr><th></th><th>西直门</th><th>通州</th><th>小计</th></tr></thead>
    <tbody>
      ${row('门诊会诊·200元档', fmtMoney(op.xizhimen_high?.amount), fmtMoney(op.tongzhou_high?.amount))}
      ${row('门诊会诊·20元档', fmtMoney(op.xizhimen_normal?.amount), fmtMoney(op.tongzhou_normal?.amount))}
      <tr><td colspan="3"><strong>门诊会诊合计</strong></td><td class="num"><strong>${fmtMoney(op.subtotal)}</strong></td></tr>
      ${row('住院会诊·200元档', fmtMoney(ip.xizhimen_high?.amount), fmtMoney(ip.tongzhou_high?.amount))}
      ${row('住院会诊·20元档', fmtMoney(ip.xizhimen_normal?.amount), fmtMoney(ip.tongzhou_normal?.amount))}
      <tr><td colspan="3"><strong>住院会诊合计</strong></td><td class="num"><strong>${fmtMoney(ip.subtotal)}</strong></td></tr>
      <tr><td><strong>共计</strong></td><td class="num"><strong>${fmtMoney(c.total?.xizhimen)}</strong></td><td class="num"><strong>${fmtMoney(c.total?.tongzhou)}</strong></td><td class="num"><strong>${fmtMoney(c.total?.grand)}</strong></td></tr>
    </tbody>
  </table>`;
}

function chartScripts(data) {
  const w = data.workload;
  const r = data.revenue;
  const c = data.consultation;

  const names = (w?.items || []).map(i => i.name);
  const cases = (w?.items || []).map(i => i.cases ?? 0);
  const blocks = (w?.items || []).map(i => i.blocks ?? 0);
  const fzCases = (w?.items || []).map(i => i.frozen_cases ?? 0);
  const fzBlocks = (w?.items || []).map(i => i.frozen_blocks ?? 0);

  const revLabels = ['病房', '门诊', '会诊(门诊)', '会诊(住院)'];
  const revTotal = [
    r?.inpatient?.total ?? 0,
    r?.outpatient?.total ?? 0,
    r?.consultOutpatient?.total ?? 0,
    r?.consultInpatient?.total ?? 0,
  ];
  const revTz = [
    r?.inpatient?.tongzhou ?? 0,
    r?.outpatient?.tongzhou ?? 0,
    r?.consultOutpatient?.tongzhou ?? 0,
    r?.consultInpatient?.tongzhou ?? 0,
  ];

  const consCat = ['门诊会诊', '住院会诊'];
  const consXz = [
    (c?.outpatient?.xizhimen_high?.amount || 0) + (c?.outpatient?.xizhimen_normal?.amount || 0),
    (c?.inpatient?.xizhimen_high?.amount || 0) + (c?.inpatient?.xizhimen_normal?.amount || 0),
  ];
  const consTz = [
    (c?.outpatient?.tongzhou_high?.amount || 0) + (c?.outpatient?.tongzhou_normal?.amount || 0),
    (c?.inpatient?.tongzhou_high?.amount || 0) + (c?.inpatient?.tongzhou_normal?.amount || 0),
  ];

  const detailLabels = (r?.outpatientDetails || []).map(d => `${d.price}元`);
  const detailTotal = (r?.outpatientDetails || []).map(d => d.total?.amount ?? 0);
  const detailTz = (r?.outpatientDetails || []).map(d => d.tongzhou?.amount ?? 0);

  const json = JSON.stringify({
    names,
    cases,
    blocks,
    fzCases,
    fzBlocks,
    revLabels,
    revTotal,
    revTz,
    consCat,
    consXz,
    consTz,
    grandXz: c?.total?.xizhimen ?? 0,
    grandTz: c?.total?.tongzhou ?? 0,
    detailLabels,
    detailTotal,
    detailTz,
  });

  return `<script src="https://cdn.staticfile.net/echarts/5.5.0/echarts.min.js"></script>
<script>
(function(){
var D=${json};
function fmtY(v){
  if(v==null) return '';
  if(v===0) return '0';
  var n=Math.abs(Number(v));
  if(n>=100000000) return (Number(v)/100000000).toFixed(1)+'亿';
  if(n>=10000) return Math.round(Number(v)/10000)+'万';
  return String(v);
}
function fmtTip(v){
  if(v==null) return '—';
  return Number(v).toLocaleString('zh-CN');
}
function moneyTooltip(){
  return { trigger:'axis', valueFormatter: fmtTip };
}
function chartBase(opts){
  opts = opts || {};
  var bottom = opts.bottom || 60;
  var left = opts.left || 56;
  return {
    tooltip: opts.tooltip || {trigger:'axis'},
    grid:{left:left,right:24,bottom:bottom,top:opts.top||72,containLabel:true},
    legend:{top:32},
    title:{left:'center',top:6,textStyle:{fontSize:14,fontWeight:600}},
    textStyle:{fontFamily:'"Noto Sans SC","Inter",sans-serif'}
  };
}
function moneyAxis(name){
  return { type:'value', name:name||'元', axisLabel:{formatter:fmtY}, nameTextStyle:{padding:[0,0,0,-10]} };
}
function go(){
  if(typeof echarts==='undefined')return;
  var rotateW = D.names.length>5?28:0;
  var rotateD = D.detailLabels.length>8?35:0;
  var bMore = D.names.length>5?72:54;
  var bMoreD = D.detailLabels.length>8?70:54;
  var el=document.getElementById('wl-chart-cases');
  if(el&&D.names.length) echarts.init(el).setOption(Object.assign(chartBase({bottom:bMore,left:54}),{
    title:{text:'工作量·例数对比',left:'center',top:6,textStyle:{fontSize:14,fontWeight:600}},
    tooltip:{trigger:'axis',valueFormatter:fmtTip},
    xAxis:{type:'category',data:D.names,axisLabel:{rotate:rotateW,interval:0}},
    yAxis:{type:'value',name:'例'},
    series:[{type:'bar',name:'例数',data:D.cases,itemStyle:{color:'#2563eb'},barMaxWidth:42,label:{show:true,position:'top',fontSize:11,formatter:function(p){return p.value?Number(p.value).toLocaleString('zh-CN'):'';}}}]
  }));
  el=document.getElementById('wl-chart-metrics');
  if(el&&D.names.length){
    var sub=D.names.map(function(_,i){ return (D.blocks[i]||0)+(D.fzCases[i]||0)+(D.fzBlocks[i]||0); });
    var any=sub.some(function(v){return v>0;});
    if(any) echarts.init(el).setOption(Object.assign(chartBase({bottom:bMore,left:64}),{
      title:{text:'工作量·多指标（块 / 冰冻例 / 冰冻块）',left:'center',top:6,textStyle:{fontSize:14,fontWeight:600}},
      tooltip:{trigger:'axis',valueFormatter:fmtTip},
      xAxis:{type:'category',data:D.names,axisLabel:{rotate:rotateW,interval:0}},
      yAxis:{type:'value',axisLabel:{formatter:fmtY}},
      series:[
        {type:'bar',name:'块数',data:D.blocks,itemStyle:{color:'#0d9488'}},
        {type:'bar',name:'冰冻例',data:D.fzCases,itemStyle:{color:'#d97706'}},
        {type:'bar',name:'冰冻块',data:D.fzBlocks,itemStyle:{color:'#7c3aed'}}
      ]
    }));
  }
  el=document.getElementById('wl-chart-revenue');
  if(el&&D.revTotal.some(function(v){return v>0;})) echarts.init(el).setOption(Object.assign(chartBase({bottom:64,left:64}),{
    title:{text:'收入金额·全额 vs 通州',left:'center',top:6,textStyle:{fontSize:14,fontWeight:600}},
    tooltip:moneyTooltip(),
    xAxis:{type:'category',data:D.revLabels,axisLabel:{interval:0}},
    yAxis:moneyAxis('元'),
    series:[
      {type:'bar',name:'全额',data:D.revTotal,itemStyle:{color:'#0369a1'},barMaxWidth:48,label:{show:true,position:'top',fontSize:10,formatter:function(p){return p.value?fmtY(p.value):'';}}},
      {type:'bar',name:'通州',data:D.revTz,itemStyle:{color:'#059669'},barMaxWidth:48,label:{show:true,position:'top',fontSize:10,formatter:function(p){return p.value?fmtY(p.value):'';}}}
    ]
  }));
  el=document.getElementById('wl-chart-rev-detail');
  if(el&&D.detailLabels.length) echarts.init(el).setOption(Object.assign(chartBase({bottom:bMoreD,left:64}),{
    title:{text:'门诊具体档位·金额对比',left:'center',top:6,textStyle:{fontSize:14,fontWeight:600}},
    tooltip:moneyTooltip(),
    xAxis:{type:'category',data:D.detailLabels,axisLabel:{interval:0,rotate:rotateD}},
    yAxis:moneyAxis('元'),
    series:[
      {type:'bar',name:'全额',data:D.detailTotal,itemStyle:{color:'#7c3a2d'}},
      {type:'bar',name:'通州',data:D.detailTz,itemStyle:{color:'#ca8a04'}}
    ]
  }));
  el=document.getElementById('wl-chart-consult');
  if(el&&(D.consXz.some(function(v){return v>0})||D.consTz.some(function(v){return v>0}))) echarts.init(el).setOption(Object.assign(chartBase({bottom:54,left:64}),{
    title:{text:'会诊统计·西直门 vs 通州',left:'center',top:6,textStyle:{fontSize:14,fontWeight:600}},
    tooltip:moneyTooltip(),
    xAxis:{type:'category',data:D.consCat},
    yAxis:moneyAxis('元'),
    series:[
      {type:'bar',name:'西直门',data:D.consXz,itemStyle:{color:'#4f46e5'},barMaxWidth:60,label:{show:true,position:'top',fontSize:11,formatter:function(p){return p.value?fmtY(p.value):'';}}},
      {type:'bar',name:'通州',data:D.consTz,itemStyle:{color:'#db2777'},barMaxWidth:60,label:{show:true,position:'top',fontSize:11,formatter:function(p){return p.value?fmtY(p.value):'';}}}
    ]
  }));
  el=document.getElementById('wl-chart-consult-grand');
  if(el&&(D.grandXz>0||D.grandTz>0)) echarts.init(el).setOption({
    title:{text:'会诊·院区合计',left:'center',top:6,textStyle:{fontSize:14,fontWeight:600}},
    tooltip:{trigger:'item',valueFormatter:fmtTip},
    grid:{left:64,right:24,bottom:48,top:54,containLabel:true},
    xAxis:{type:'category',data:['西直门','通州']},
    yAxis:moneyAxis('元'),
    series:[{type:'bar',data:[D.grandXz,D.grandTz],itemStyle:{color:function(p){return p.dataIndex===0?'#4f46e5':'#db2777';}},barMaxWidth:64,label:{show:true,position:'top',fontSize:12,formatter:function(p){return p.value?fmtY(p.value):'';}}}]
  });
}
if(document.readyState==='loading') document.addEventListener('DOMContentLoaded',go); else go();
window.addEventListener('resize',function(){
  document.querySelectorAll('[id^="wl-chart-"]').forEach(function(dom){
    var ch=echarts.getInstanceByDom(dom); if(ch) ch.resize();
  });
});
})();
</script>`;
}

function downloadButtonsRow(chartId) {
  return `<div class="wl-actions"><button type="button" class="wl-dl-chart" data-chart="${esc(chartId)}">下载该图为 PNG</button></div>`;
}

function workloadBodyInner(data, withDlButtons) {
  const period = data.meta?.period || data.workload?.period || '';
  const dl = id => (withDlButtons ? downloadButtonsRow(id) : '');

  return `
<div class="wl-report">
  <h1 class="wl-h1">病理科月度工作报表·可视化</h1>
  <p class="wl-meta">${esc(period) || '（请在模板第一行标题中填写年月，如「2026年4月病理科工作量」）'}</p>

  <section class="wl-section">
    <h2>一、工作量统计</h2>
    ${workloadTableHTML(data)}
    <div id="wl-chart-cases" class="wl-chart wl-chart-sm"></div>
    ${dl('wl-chart-cases')}
    <div id="wl-chart-metrics" class="wl-chart wl-chart-sm"></div>
    ${dl('wl-chart-metrics')}
  </section>

  <section class="wl-section">
    <h2>二、工作量与收入金额</h2>
    ${revenueTableHTML(data)}
    <div id="wl-chart-revenue" class="wl-chart"></div>
    ${dl('wl-chart-revenue')}
    ${(data.revenue?.outpatientDetails?.length ? `<div id="wl-chart-rev-detail" class="wl-chart wl-chart-sm"></div>${dl('wl-chart-rev-detail')}` : '')}
  </section>

  <section class="wl-section">
    <h2>三、会诊统计</h2>
    ${consultationTableHTML(data)}
    <div id="wl-chart-consult" class="wl-chart wl-chart-sm"></div>
    ${dl('wl-chart-consult')}
    <div id="wl-chart-consult-grand" class="wl-chart-sm wl-chart"></div>
    ${dl('wl-chart-consult-grand')}
  </section>
</div>`;
}

function generateWorkloadReportBody(data, opts = {}) {
  const showDl = opts.downloadButtons !== false;
  const inner = workloadBodyInner(data, showDl);
  return workloadCSS() + inner + chartScripts(data);
}

function generateWorkloadReport(data) {
  return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>病理科工作报表_${esc(data.meta?.period || '')}</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&family=Noto+Sans+SC:wght@400;600;700&display=swap" rel="stylesheet">
${workloadCSS()}
</head>
<body>
${workloadBodyInner(data, false)}
${chartScripts(data)}
</body></html>`;
}

module.exports = { generateWorkloadReport, generateWorkloadReportBody, workloadCSS };
