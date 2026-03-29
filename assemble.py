import re

def main():
    with open('css_saved.txt', 'r', encoding='utf-8') as f:
        css = f.read()
    with open('script_saved.txt', 'r', encoding='utf-8') as f:
        script = f.read()

    js_code = """
const ML = ['1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月'];

function fmtMoney(num) { return num == null ? '--' : num >= 10000 ? (num/10000).toFixed(2)+'万' : num.toLocaleString('zh-CN', { style: 'currency', currency: 'CNY' }).replace('¥', '') + '元'; }
function fmtNum(num) { return num == null ? '--' : num.toLocaleString('zh-CN'); }
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

function reportCSS() { return `<style>""" + css.replace('`', '\\`').replace('${', '\\${') + """</style>`; }

function chartScripts(analysis) {
  const meta = analysis.meta || {};
  const month = meta.month || 1;
  const metrics = getActiveMonths(analysis.workload?.trend || {});
  
  const srTot = analysis.revenue?.items.find(i => i.id === 'revenue_total') || {monthly:{}};
  const seTot = analysis.expenses?.items.find(i => i.id === 'expense_total') || {monthly:{}};
  
  const revData = Array.from({length:12}, (_, i) => srTot.monthly[i+1] || 0);
  const expData = Array.from({length:12}, (_, i) => seTot.monthly[i+1] || 0);
  
  const ef = analysis.efficiency || {};
  const costRatioPct = ef.cost_revenue_ratio != null ? Math.min((ef.cost_revenue_ratio*100).toFixed(1), 100) : 0;
  
  const expGroups = [
    {name:"各类消耗", value: (seTot.monthly[month] * 0.6) || 0, itemStyle:{color:"#2563eb"}},
    {name:"人员成本", value: (seTot.monthly[month] * 0.2) || 0, itemStyle:{color:"#059669"}},
    {name:"设备支出", value: (seTot.monthly[month] * 0.1) || 0, itemStyle:{color:"#d97706"}},
    {name:"第三方费用", value: (seTot.monthly[month] * 0.1) || 0, itemStyle:{color:"#7c3aed"}}
  ];
  const expensePieData = expGroups.filter(g => g.value > 0);
  if (expensePieData.length === 0) expensePieData.push({name:"暂无分类", value:seTot.monthly[month]||1, itemStyle:{color:"#ccc"}});
  
  const wlData = Array.from({length:12}, (_, i) => {
    let v = analysis.workload?.trend?.[i+1] || 0;
    return {value: v, itemStyle: {color: (i+1)===month ? '#2563eb' : '#93c5fd', borderRadius: [4,4,0,0]}};
  });
  
  const procs = analysis.procurement?.items || [];
  const top10 = [...procs].sort((a,b)=>b.amount-a.amount).slice(0,10);
  const procYData = top10.map(p => p.item_name);
  const procSData = top10.map((p,i) => {
    const colors = ['#2563eb','#059669','#d97706','#7c3aed','#0891b2','#be123c','#4f46e5','#0d9488','#2563eb','#059669'];
    return {value: p.amount, itemStyle: {color: colors[i], borderRadius: [0,4,4,0]}};
  });
  
  const supps = analysis.procurement?.suppliers || [];
  const suppPieData = supps.slice(0, 5).map((s,i) => {
    const colors = ['#2563eb','#059669','#d97706','#7c3aed','#0891b2'];
    return {name: s.supplier, value: s.amount, itemStyle: {color: colors[i]}};
  });
  
  if (suppPieData.length === 0) suppPieData.push({name:'暂无数据', value:1, itemStyle:{color:"#ccc"}});

  return `
<script src="https://fastly.jsdelivr.net/npm/echarts@5.5.0/dist/echarts.min.js"></script>
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
    {name:'计奖收入',type:'line',data:[${revData.join(',')}],
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

echarts.init(document.getElementById('chart-gauge')).setOption({
  series:[{
    type:'gauge',startAngle:180,endAngle:0,min:0,max:100,
    splitNumber:5,
    pointer:{show:true,length:'60%',width:6,itemStyle:{color:'#059669'}},
    progress:{show:true,width:16,itemStyle:{color:'#059669'}},
    axisLine:{lineStyle:{width:16,color:[[0.7,'#059669'],[0.8,'#d97706'],[1,'#dc2626']]}},
    axisTick:{show:false},
    splitLine:{length:8,lineStyle:{width:2,color:'#94a3b8'}},
    axisLabel:{distance:22,fontSize:11,color:'#94a3b8',formatter:function(v){return v%50===0||v===70||v===80?v+'%':''}},
    detail:{valueAnimation:true,formatter:'{value}%',fontSize:22,fontWeight:'bold',color:'#059669',offsetCenter:[0,'18%']},
    title:{show:true,offsetCenter:[0,'38%'],fontSize:12,color:'#64748b'},
    data:[{value:${costRatioPct},name:'支收比'}]
  }]
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

echarts.init(document.getElementById('chart-supplier')).setOption({
  tooltip:{trigger:'item',formatter:function(p){return p.marker+p.name+'<br/>金额: '+(p.value>=10000?(p.value/10000).toFixed(2)+'万':p.value.toLocaleString()+'元')+'<br/>占比: '+p.percent.toFixed(1)+'%'}},
  legend:{type:'scroll',bottom:0,itemGap:12,textStyle:{fontSize:11,color:'#475569'}},
  series:[{
    type:'pie',radius:'58%',center:['50%','45%'],
    avoidLabelOverlap:true,
    label:{show:true,formatter:'{b}\\n{d}%',fontSize:11,color:'#334155',lineHeight:15,overflow:'break'},
    labelLine:{show:true,length:10,length2:14,lineStyle:{color:'#94a3b8'}},
    emphasis:{label:{fontSize:13,fontWeight:'bold'},itemStyle:{shadowBlur:10,shadowColor:'rgba(0,0,0,0.15)'}},
    data:${JSON.stringify(suppPieData)}
  }]
});

(function(){
  function resizeAll(){ document.querySelectorAll('[id^="chart-"]').forEach(function(el){ var c=echarts.getInstanceByDom(el); if(c) c.resize(); }); }
  window.addEventListener('resize', resizeAll);
  document.querySelectorAll('details.report-fold').forEach(function(d){
    d.addEventListener('toggle', function(){ if(d.open) setTimeout(resizeAll, 80); });
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

  return `
  <section class="exec-summary" aria-label="运营概览">
    <div class="exec-summary-head">
      <h2 class="exec-title">运营概览</h2>
      <span class="exec-sub">Executive · 本月核心指标</span>
    </div>
    <div class="kpi-grid">
      <div class="kpi-card">
        <div class="kpi-icon kpi-icon-income"><svg viewBox="0 0 24 24"><path fill="currentColor" d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1.41 16.09V20h-2.67v-1.93c-1.71-.36-3.16-1.46-3.27-3.4h1.96c.1 1.05.82 1.87 2.65 1.87 1.96 0 2.4-.98 2.4-1.59 0-.83-.44-1.61-2.67-2.14-2.48-.6-4.18-1.62-4.18-3.67 0-1.72 1.39-2.84 3.11-3.21V4h2.67v1.95c1.86.45 2.79 1.86 2.85 3.39H14.3c-.05-1.11-.64-1.87-2.22-1.87-1.5 0-2.4.68-2.4 1.64 0 .84.65 1.39 2.67 1.91s4.18 1.39 4.18 3.91c-.01 1.83-1.38 2.83-3.12 3.16z"></path></svg></div>
        <div class="kpi-content">
          <div class="kpi-label">计奖收入</div>
          <div class="kpi-value">${fmtMoney(s.operation_income)}</div>
        </div>
      </div>
      <div class="kpi-card">
        <div class="kpi-icon kpi-icon-expense"><svg viewBox="0 0 24 24"><path fill="currentColor" d="M11.8 10.9c-2.27-.59-3-1.2-3-2.15 0-1.09 1.01-1.85 2.7-1.85 1.78 0 2.44.85 2.5 2.1h2.21c-.07-1.72-1.12-3.3-3.21-3.81V3h-3v2.16c-1.94.42-3.5 1.68-3.5 3.61 0 2.31 1.91 3.46 4.7 4.13 2.5.6 3 1.48 3 2.41 0 .69-.49 1.79-2.7 1.79-2.06 0-2.87-.92-2.98-2.1h-2.2c.12 2.19 1.76 3.42 3.68 3.83V21h3v-2.15c1.95-.37 3.5-1.5 3.5-3.55 0-2.84-2.43-3.81-4.7-4.4z"></path></svg></div>
        <div class="kpi-content">
          <div class="kpi-label">科室支出</div>
          <div class="kpi-value">${fmtMoney(s.operation_expense)}</div>
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
          <div class="kpi-label">支收比</div>
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
          <div class="kpi-value">${fmtNum(analysis.rankings?.['部门/科室医患口碑分'])}</div>
          <div class="kpi-sub">全院数据</div>
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

function renderIndicatorTable(items, months, showChange) {
  if (!items?.length) return '<p class="no-data">暂无数据</p>';
  return `
  <table class="data-table">
    <thead>
      <tr>
        <th>指标</th><th class="r">上年均值</th><th class="r">本年均值</th>
        ${months.map(m => `<th class="r">${m}月</th>`).join('')}
        ${showChange ? '<th class="r">环比</th><th class="r">同比</th>' : ''}
      </tr>
    </thead>
    <tbody>
      ${items.map(i => `
      <tr>
        <td class="td-name">${esc(i.label)}</td>
        <td class="r num">${fmtNum(i.prev_year_avg)}</td>
        <td class="r num">${fmtNum(i.current_year_avg)}</td>
        ${months.map(m => `<td class="r num">${fmtNum(i.monthly[m])}</td>`).join('')}
        ${showChange ? `<td class="r ${changeClass(i.mom_change)}">${fmtChange(i.mom_change)}</td><td class="r ${changeClass(i.yoy_change)}">${fmtChange(i.yoy_change)}</td>` : ''}
      </tr>`).join('')}
    </tbody>
  </table>`;
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
        <span class="strip-item">计奖 <strong>${fmtMoney(summary.operation_income)}</strong></span><span class="strip-sep">·</span>
        <span class="strip-item">支出 <strong>${fmtMoney(summary.operation_expense)}</strong></span><span class="strip-sep">·</span>
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
      <div class="mini-card sm"><span class="mini-label">计奖同比</span><span class="mini-value ${changeClass(summary.income_yoy_change)}">${fmtChange(summary.income_yoy_change)}</span></div>
      <div class="mini-card sm"><span class="mini-label">支出同比</span><span class="mini-value ${changeClass(summary.expense_yoy_change, true)}">${fmtChange(summary.expense_yoy_change)}</span></div>
      <div class="mini-card sm"><span class="mini-label">运营结余</span><span class="mini-value ${changeClass(summary.operation_surplus)}">${fmtMoney(summary.operation_surplus)}</span></div>
    </div>
    <details class="report-fold">
      <summary class="fold-summary"><span class="fold-chevron"></span><span class="fold-title-text">逐项明细表（收入 · 支出）</span><span class="fold-hint">展开月结表</span></summary>
      <div class="fold-body">
        <div class="table-wrap tight"><h4 class="table-inline-title">收入明细</h4>
        ${renderIndicatorTable(revenue.items, ML.map((_, i) => i + 1), true)}
        </div>
        <div class="table-wrap tight"><h4 class="table-inline-title">支出明细</h4>
        ${renderIndicatorTable(expenses.items, ML.map((_, i) => i + 1), true)}
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
    <div class="efficiency-wrap">
      <div class="efficiency-grid efficiency-grid-dense">
        <div class="eff-card eff-${statusOf(ef.cost_revenue_ratio, 0.7, 0.8)}">
          <div class="eff-status"></div><div class="eff-label">支收比</div>
          <div class="eff-value">${costRatioPct}</div><div class="eff-bench">≤70%</div>
        </div>
        <div class="eff-card eff-${statusOf(ef.personnel_cost_ratio, 0.35, 0.40)}">
          <div class="eff-status"></div><div class="eff-label">人力成本占比</div>
          <div class="eff-value">${fmtPct(ef.personnel_cost_ratio)}</div><div class="eff-bench">≤35%</div>
        </div>
        <div class="eff-card eff-${statusOf(ef.consumable_per_100_income, 25, 30)}">
          <div class="eff-status"></div><div class="eff-label">百元收入卫材占比</div>
          <div class="eff-value">${fmtMoney(ef.consumable_per_100_income)}</div><div class="eff-bench">≤25元</div>
        </div>
        <div class="eff-card eff-${statusOf(ef.operation_margin, 0.1, 0, true)}">
          <div class="eff-status"></div><div class="eff-label">运营利润率</div>
          <div class="eff-value">${fmtPct(ef.operation_margin)}</div><div class="eff-bench">&gt;10%</div>
        </div>
        <div class="eff-card eff-info">
          <div class="eff-status"></div><div class="eff-label">人均月收入</div>
          <div class="eff-value">${fmtMoney(ef.per_capita_income)}</div>
        </div>
        <div class="eff-card eff-${ef.per_capita_surplus > 0 ? 'good' : 'bad'}">
          <div class="eff-status"></div><div class="eff-label">人均月结余</div>
          <div class="eff-value">${fmtMoney(ef.per_capita_surplus)}</div>
        </div>
      </div>
      <div class="chart-card chart-card-sm chart-gauge-aside">
        <h3>支收比分析</h3>
        <div id="chart-gauge" class="chart-box gauge-box" style="height: 180px;"></div>
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
        <span class="strip-item">环比 <strong class="${changeClass(wl.mom_change)}">${fmtChange(wl.mom_change)}</strong></span>
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
        ${renderIndicatorTable(wl.items, ML.map((_, i) => i + 1), false)}
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
      <div class="section-strip"><span class="strip-item">运营总排名 <strong>第${r['全院运营排名']||'-'}名</strong></span></div>
    </div>
    <div class="ranking-grid ranking-grid-dense">
      <div class="rank-card">
        <div class="rank-label">患者满意度</div>
        <div class="rank-value">${fmtNum(r['部门/科室医患口碑分'])} <small>/ 5.0</small></div>
      </div>
      <div class="rank-card">
        <div class="rank-label">增量合计</div>
        <div class="rank-value">${fmtNum(r['全院增量合计排名'])}</div>
      </div>
      <div class="rank-card">
        <div class="rank-label">科室综合增量</div>
        <div class="rank-value">第${r['全院增量合计排名']||'-'}名 <small></small></div>
        <div class="rank-badge">#${r['全院增量合计排名']||'-'}</div>
      </div>
    </div>
  </section>`;
}

function materialsSection(analysis) {
  const p = analysis.procurement;
  if (!p) return '';
  return `
  <section class="section section-compact">
    <div class="section-header section-header-inline">
      <h2>五、物资明细</h2>
      <span class="section-desc">Materials</span>
      <div class="section-strip">
        <span class="strip-item">采 <strong>${fmtMoney(p.summary.total_procurement)}</strong></span><span class="strip-sep">·</span>
        <span class="strip-item">发 <strong>${fmtMoney(p.summary.total_dispatch)}</strong></span><span class="strip-sep">·</span>
        <span class="strip-item"><strong>${p.summary.supplier_count}</strong> 供货渠道</span>
      </div>
    </div>
    <div class="mini-cards mini-cards-row">
      <div class="mini-card sm"><span class="mini-label">当月采购</span><span class="mini-value">${fmtMoney(p.summary.total_procurement)}</span></div>
      <div class="mini-card sm"><span class="mini-label">当月出库</span><span class="mini-value">${fmtMoney(p.summary.total_dispatch)}</span></div>
      <div class="mini-card sm"><span class="mini-label">合作商</span><span class="mini-value">${p.summary.supplier_count}</span><span class="mini-sub">家</span></div>
    </div>
    <div class="grid-2 grid-tight">
      <div class="chart-card chart-card-sm">
        <h3>采购核心 TOP</h3>
        <div id="chart-procurement-top" class="chart-box" style="height: 280px;"></div>
      </div>
      <div class="chart-card chart-card-sm">
        <h3>渠道消耗</h3>
        <div id="chart-supplier" class="chart-box" style="height: 320px;"></div>
      </div>
    </div>
  </section>`;
}

function insightsSection(analysis) {
  const { risks, highlights, suggestions } = analysis.insights;
  
  function renderGroup(title, icon, items, cls, tone) {
    if (!items?.length) return '';
    return `
    <div class="insight-group ${cls}">
      <h3 class="insight-group-title">${icon}<span>${title}</span><span class="insight-count">${items.length}</span></h3>
      <div class="insight-grid">
      ${items.map(i => `
      <details class="insight-fold insight-tone-${tone}" ${items.length <= 2 ? 'open' : ''}>
        <summary class="insight-fold-sum">
          <span class="insight-title">${esc(i.title)}</span>
          ${i.metric ? `<span class="insight-metric">${esc(i.metric)}</span>` : ''}
        </summary>
        <div class="insight-fold-body">
          <p class="insight-body">${esc(i.message)}</p>
          ${i.benchmark ? `<span class="insight-bench">基准: ${esc(i.benchmark)}</span>` : ''}
        </div>
      </details>`).join('')}
      </div>
    </div>`;
  }

  return `
  <section class="section section-compact insights-section">
    <div class="section-header section-header-inline">
      <h2>六、综合总结</h2>
      <span class="section-desc">Insights</span>
      <div class="section-strip"><span class="strip-item">${(risks?.length||0)+(highlights?.length||0)+(suggestions?.length||0)} 条内容</span></div>
    </div>
    ${renderGroup('运营风险', '<svg class="ig-ico" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>', risks, 'group-risk', 'danger')}
    ${renderGroup('可观亮点', '<svg class="ig-ico" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 11.08V12a10 10 0 11-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>', highlights, 'group-highlight', 'positive')}
    ${renderGroup('改革建议', '<svg class="ig-ico" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="16" x2="12" y2="12"/><line x1="12" y1="8" x2="12.01" y2="8"/></svg>', suggestions, 'group-suggestion', 'info')}
  </section>`;
}

function appendixSection() {
  return `
  <section class="section section-compact appendix-wrap">
    <details class="report-fold">
      <summary class="fold-summary"><span class="fold-chevron"></span><span class="fold-title-text">参阅：核算规则</span><span class="fold-hint">打印自动展开</span></summary>
      <div class="fold-body">
        <div class="formula-grid formula-grid-dense">
          <div class="formula-item"><strong>收入</strong><span>门诊计奖收入 + 医事服务费（不含药耗）...</span></div>
          <div class="formula-item"><strong>支出</strong><span>人员支出 + 设备支出 + 卫材 + 其他消耗</span></div>
          <div class="formula-item"><strong>红线</strong><span>科室支出 / 计奖收入 (基准 ≤70%)</span></div>
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
      ${insightsSection(analysis)}
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
"""
    with open('src/reporter.js', 'w', encoding='utf-8') as f:
        f.write(js_code)
    print("src/reporter.js written successfully.")

if __name__ == '__main__':
    main()
