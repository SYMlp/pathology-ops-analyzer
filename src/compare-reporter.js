function fmtAmount(n) {
  if (n == null) return '—';
  if (Math.abs(n) >= 10000) return (n / 10000).toFixed(1) + ' 万';
  return n.toLocaleString();
}
function fmtNum(n) {
  if (n == null) return '—';
  return n.toLocaleString();
}
function fmtPct(n, digits = 1) {
  if (n == null || !isFinite(n)) return '—';
  return n.toFixed(digits) + '%';
}
function totalChange(arr) {
  const first = arr.find(x => x != null);
  const last = [...arr].reverse().find(x => x != null);
  if (first == null || last == null || first === 0) return null;
  return (last - first) / first * 100;
}

function buildSeries(monthly) {
  const months = monthly.map(m => m.meta.period);
  const monthsShort = monthly.map(m => `${m.meta.month}月`);

  const W = monthly.map(m => m.workload || {});
  const C = monthly.map(m => m.consultation || {});
  const R = monthly.map(m => m.revenue || {});

  const grandTotal = R.map(r => r.grand_total);
  const grandTongzhou = R.map(r => r.grand_tongzhou);
  const grandXizhimen = R.map(r => (r.grand_total || 0) - (r.grand_tongzhou || 0));
  const tongzhouRatio = R.map(r => r.grand_total ? (r.grand_tongzhou / r.grand_total * 100) : null);

  const inpatientRev = R.map(r => r.inpatient_total);
  const outpatientRev = R.map(r => r.outpatient_total);
  const consultORev = R.map(r => r.consult_outpatient_total);
  const consultIRev = R.map(r => r.consult_inpatient_total);

  const waijianCases = W.map(w => w.waijian_cases);
  const waijianBlocks = W.map(w => w.waijian_blocks);
  const tzWaijianCases = W.map(w => w.tongzhou_waijian_cases);
  const tzWaijianBlocks = W.map(w => w.tongzhou_waijian_blocks);
  const frozenCases = W.map(w => w.frozen_cases);

  const fukeReview = W.map(w => w.fuke_review);
  const shenneiReview = W.map(w => w.shennei_review);
  const genetic = W.map(w => w.genetic);

  const consultOAmt = C.map(c => c.outpatient_amount);
  const consultIAmt = C.map(c => c.inpatient_amount);
  const consultTotalAmt = C.map(c => c.consult_amount_total);
  const consultTotalQty = C.map(c => c.consult_qty_total);

  return {
    months, monthsShort,
    grandTotal, grandTongzhou, grandXizhimen, tongzhouRatio,
    inpatientRev, outpatientRev, consultORev, consultIRev,
    waijianCases, waijianBlocks, tzWaijianCases, tzWaijianBlocks, frozenCases,
    fukeReview, shenneiReview, genetic,
    consultOAmt, consultIAmt, consultTotalAmt, consultTotalQty,
  };
}

function compareCSS() {
  return `
.compare-scope {
  --c-bg: #f8fafc;
  --c-card: #ffffff;
  --c-border: #e2e8f0;
  --c-text: #0f172a;
  --c-muted: #64748b;
  font-family: 'PingFang SC', 'Microsoft YaHei', system-ui, -apple-system, sans-serif;
  color: var(--c-text);
  line-height: 1.5;
  background: var(--c-bg);
}
.compare-scope * { box-sizing: border-box; }
.compare-scope .cmp-container { max-width: 1280px; margin: 0 auto; padding: 24px 28px 48px; }

.compare-scope .cmp-report-header {
  background: linear-gradient(135deg, #1e3a8a 0%, #3b82f6 100%);
  color: #fff;
  padding: 28px 32px;
  border-radius: 16px;
  margin-bottom: 24px;
  box-shadow: 0 4px 12px rgba(59, 130, 246, 0.15);
}
.compare-scope .cmp-report-header h1 { margin: 0 0 8px; font-size: 26px; font-weight: 700; color: #fff; }
.compare-scope .cmp-report-header .cmp-meta { opacity: 0.9; font-size: 14px; }
.compare-scope .cmp-report-header .cmp-badges { margin-top: 14px; display: flex; gap: 8px; flex-wrap: wrap; }
.compare-scope .cmp-report-header .cmp-badge {
  background: rgba(255, 255, 255, 0.18);
  padding: 4px 10px;
  border-radius: 999px;
  font-size: 12px;
}

.compare-scope .cmp-section-title {
  font-size: 13px;
  color: var(--c-muted);
  text-transform: uppercase;
  letter-spacing: 0.06em;
  margin: 24px 0 12px;
  font-weight: 600;
}

.compare-scope .cmp-kpi-grid {
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: 14px;
  margin-bottom: 8px;
}
@media (max-width: 1080px) { .compare-scope .cmp-kpi-grid { grid-template-columns: repeat(2, 1fr); } }
@media (max-width: 640px)  { .compare-scope .cmp-kpi-grid { grid-template-columns: 1fr; } }

.compare-scope .cmp-kpi-card {
  background: var(--c-card);
  border: 1px solid var(--c-border);
  border-radius: 12px;
  padding: 12px 14px 8px;
  display: flex;
  flex-direction: column;
}
.compare-scope .cmp-kpi-head { display: flex; justify-content: space-between; align-items: baseline; }
.compare-scope .cmp-kpi-label { font-size: 13px; font-weight: 600; color: var(--c-text); }
.compare-scope .cmp-kpi-total { font-size: 13px; font-weight: 700; }
.compare-scope .cmp-kpi-spark { height: 100px; margin: 4px -2px 0; }

.compare-scope .cmp-chart-card {
  background: var(--c-card);
  border: 1px solid var(--c-border);
  border-radius: 12px;
  padding: 18px 20px 16px;
  margin-bottom: 16px;
}
.compare-scope .cmp-chart-card-head {
  display: flex;
  align-items: baseline;
  justify-content: space-between;
  margin-bottom: 8px;
  gap: 12px;
  flex-wrap: wrap;
}
.compare-scope .cmp-chart-title { font-size: 17px; font-weight: 700; margin: 0; color: var(--c-text); }
.compare-scope .cmp-chart-subtitle { font-size: 13px; color: var(--c-muted); margin: 0; }
.compare-scope .cmp-chart-caption {
  font-size: 12px;
  color: var(--c-muted);
  margin-top: 6px;
  padding-top: 8px;
  border-top: 1px dashed var(--c-border);
}
.compare-scope .cmp-chart-canvas { width: 100%; height: 380px; }
.compare-scope .cmp-chart-canvas-tall { height: 440px; }

.compare-scope .cmp-small-grid {
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: 16px;
  margin-top: 4px;
}
@media (max-width: 900px) {
  .compare-scope .cmp-small-grid { grid-template-columns: 1fr; }
}
.compare-scope .cmp-small-cell { display: flex; flex-direction: column; }
.compare-scope .cmp-small-title {
  font-size: 14px;
  font-weight: 600;
  color: var(--c-text);
  text-align: center;
  margin: 4px 0 2px;
}
.compare-scope .cmp-small-canvas { width: 100%; height: 280px; }

.compare-scope .cmp-footer { text-align: center; color: var(--c-muted); font-size: 12px; margin-top: 24px; }
`;
}

function compareBodyInner(monthly) {
  const S = buildSeries(monthly);
  const period = `${monthly[0].meta.period} ~ ${monthly[monthly.length - 1].meta.period}`;

  const kpis = [
    { label: '总收入', values: S.grandTotal, fmt: 'amount', color: '#3b82f6' },
    { label: '外检例数', values: S.waijianCases, fmt: 'num', color: '#8b5cf6' },
    { label: '外检块数', values: S.waijianBlocks, fmt: 'num', color: '#06b6d4' },
    { label: '会诊总金额', values: S.consultTotalAmt, fmt: 'amount', color: '#22c55e' },
    { label: '通州收入占比', values: S.tongzhouRatio, fmt: 'pct', color: '#f59e0b' },
    { label: '基因检测', values: S.genetic, fmt: 'num', color: '#ef4444' },
  ];

  const kpiCards = kpis.map((k, idx) => {
    const tot = totalChange(k.values);
    const totColor = tot == null ? '#94a3b8' : (tot >= 0 ? '#16a34a' : '#dc2626');
    const totArrow = tot == null ? '' : (tot >= 0 ? '↑' : '↓');
    return `
      <div class="cmp-kpi-card">
        <div class="cmp-kpi-head">
          <span class="cmp-kpi-label">${k.label}</span>
          <span class="cmp-kpi-total" style="color:${totColor}">${tot == null ? '—' : (tot >= 0 ? '+' : '') + tot.toFixed(1) + '%'} ${totArrow}</span>
        </div>
        <div class="cmp-kpi-spark" id="cmp-spark-${idx}"></div>
      </div>`;
  }).join('');

  const lastIdx = S.tongzhouRatio.length - 1;
  const totalRevPct = totalChange(S.grandTotal);
  const totalCasePct = totalChange(S.waijianCases);

  return `
<div class="compare-scope">
  <div class="cmp-container">
    <header class="cmp-report-header">
      <h1>病理科月度工作量对比报告</h1>
      <div class="cmp-meta">${period} · 共 ${monthly.length} 个月</div>
      <div class="cmp-badges">
        <span class="cmp-badge">总收入 末月 vs 首月 ${totalRevPct == null ? '—' : (totalRevPct >= 0 ? '+' : '') + totalRevPct.toFixed(1) + '%'}</span>
        <span class="cmp-badge">外检例数 末月 vs 首月 ${totalCasePct == null ? '—' : (totalCasePct >= 0 ? '+' : '') + totalCasePct.toFixed(1) + '%'}</span>
        <span class="cmp-badge">通州占比 ${fmtPct(S.tongzhouRatio[0], 1)} → ${fmtPct(S.tongzhouRatio[lastIdx], 1)}</span>
      </div>
    </header>

    <div class="cmp-section-title">月度趋势小多图（一眼看 6 个核心指标的形状）</div>
    <div class="cmp-kpi-grid">${kpiCards}</div>

    <div class="cmp-section-title">月度对比详图</div>

    <div class="cmp-chart-card">
      <div class="cmp-chart-card-head">
        <div>
          <h2 class="cmp-chart-title">① 收入大盘 · 月度对比</h2>
          <p class="cmp-chart-subtitle">病房 / 门诊 / 会诊门诊 / 会诊住院 堆叠 + 总收入折线</p>
        </div>
      </div>
      <div id="cmp-chart1" class="cmp-chart-canvas cmp-chart-canvas-tall"></div>
      <div class="cmp-chart-caption">柱顶 % 标签 = 总收入环比变化；虚线 = 月份均值</div>
    </div>

    <div class="cmp-chart-card">
      <div class="cmp-chart-card-head">
        <div>
          <h2 class="cmp-chart-title">② 西直门 vs 通州 收入 · 月度对比</h2>
          <p class="cmp-chart-subtitle">两院区收入簇状柱 + 通州占比折线</p>
        </div>
      </div>
      <div id="cmp-chart2" class="cmp-chart-canvas"></div>
      <div class="cmp-chart-caption">折线右轴 = 通州收入占比 (%)；看通州院区是否在持续拉份额</div>
    </div>

    <div class="cmp-chart-card">
      <div class="cmp-chart-card-head">
        <div>
          <h2 class="cmp-chart-title">③ 外检工作量 · 月度对比</h2>
          <p class="cmp-chart-subtitle">外检例数 / 通州外检例数 / 冰冻例数 簇状柱 + 总例数折线</p>
        </div>
      </div>
      <div id="cmp-chart3" class="cmp-chart-canvas"></div>
      <div class="cmp-chart-caption">柱顶 % 标签 = 外检总例数环比；折线 = 月度趋势</div>
    </div>

    <div class="cmp-chart-card">
      <div class="cmp-chart-card-head">
        <div>
          <h2 class="cmp-chart-title">④ 会诊业务 · 月度对比</h2>
          <p class="cmp-chart-subtitle">门诊会诊 / 住院会诊 金额堆叠 + 总人次折线</p>
        </div>
      </div>
      <div id="cmp-chart4" class="cmp-chart-canvas"></div>
      <div class="cmp-chart-caption">柱顶 % 标签 = 会诊总金额环比；折线右轴 = 总人次</div>
    </div>

    <div class="cmp-chart-card">
      <div class="cmp-chart-card-head">
        <div>
          <h2 class="cmp-chart-title">⑤ 专项审核类 · 月度对比</h2>
          <p class="cmp-chart-subtitle">妇科细胞 / 肾内切片 / 基因检测 各自独立子图，柱顶 % = 月度环比</p>
        </div>
      </div>
      <div class="cmp-small-grid">
        <div class="cmp-small-cell">
          <div class="cmp-small-title" style="color:#3b82f6">妇科细胞审核</div>
          <div id="cmp-chart5a" class="cmp-small-canvas"></div>
        </div>
        <div class="cmp-small-cell">
          <div class="cmp-small-title" style="color:#8b5cf6">肾内切片审核</div>
          <div id="cmp-chart5b" class="cmp-small-canvas"></div>
        </div>
        <div class="cmp-small-cell">
          <div class="cmp-small-title" style="color:#ef4444">基因检测</div>
          <div id="cmp-chart5c" class="cmp-small-canvas"></div>
        </div>
      </div>
      <div class="cmp-chart-caption">3 个独立子图避免簇状柱视觉混淆，让每个专项的月度变化各自一目了然</div>
    </div>

    <div class="cmp-footer">病理数据平台 · 多月对比分析 · 生成于 ${new Date().toLocaleString('zh-CN')}</div>
  </div>
</div>
`;
}

function compareScripts(monthly) {
  const S = buildSeries(monthly);
  const kpis = [
    { label: '总收入', values: S.grandTotal, fmt: 'amount', color: '#3b82f6' },
    { label: '外检例数', values: S.waijianCases, fmt: 'num', color: '#8b5cf6' },
    { label: '外检块数', values: S.waijianBlocks, fmt: 'num', color: '#06b6d4' },
    { label: '会诊总金额', values: S.consultTotalAmt, fmt: 'amount', color: '#22c55e' },
    { label: '通州收入占比', values: S.tongzhouRatio, fmt: 'pct', color: '#f59e0b' },
    { label: '基因检测', values: S.genetic, fmt: 'num', color: '#ef4444' },
  ];
  const payload = { monthsShort: S.monthsShort, months: S.months, series: S, kpis };

  return `
<script>
(function(){
function initCompareCharts(){
const D = ${JSON.stringify(payload)};

const COLOR = {
  primary: '#3b82f6', tongzhou: '#f59e0b', xizhimen: '#3b82f6',
  inpatient: '#3b82f6', outpatient: '#10b981',
  consultO: '#8b5cf6', consultI: '#ec4899', freeze: '#06b6d4',
  fuke: '#3b82f6', shennei: '#8b5cf6', genetic: '#ef4444',
  up: '#16a34a', down: '#dc2626',
};

function ringPct(arr) {
  return arr.map((v, i) => {
    if (i === 0 || arr[i - 1] == null || arr[i - 1] === 0 || v == null) return null;
    return (v - arr[i - 1]) / arr[i - 1] * 100;
  });
}
function avg(arr) {
  const xs = arr.filter(x => x != null);
  if (!xs.length) return null;
  return xs.reduce((s, x) => s + x, 0) / xs.length;
}
function fmtPctSign(p) {
  if (p == null) return '';
  return (p >= 0 ? '+' : '') + p.toFixed(1) + '%';
}
function topPctLabel(values) {
  const pcts = ringPct(values);
  return {
    show: true, position: 'top', distance: 8, fontSize: 11, fontWeight: 700,
    formatter: (p) => { const pct = pcts[p.dataIndex]; return pct == null ? '' : fmtPctSign(pct); },
    color: (p) => { const pct = pcts[p.dataIndex]; return pct == null ? '#94a3b8' : (pct >= 0 ? COLOR.up : COLOR.down); },
  };
}
function avgMarkLine(values, name) {
  const a = avg(values);
  if (a == null) return undefined;
  return {
    silent: true, symbol: 'none',
    lineStyle: { type: 'dashed', color: '#94a3b8', width: 1 },
    label: { formatter: name + ': ' + (a >= 10000 ? (a / 10000).toFixed(1) + '万' : Math.round(a).toLocaleString()), color: '#64748b', fontSize: 11 },
    data: [{ yAxis: a }],
  };
}

D.kpis.forEach((k, idx) => {
  const el = document.getElementById('cmp-spark-' + idx);
  if (!el) return;
  const chart = echarts.init(el);
  const fmtFn = k.fmt === 'amount'
    ? (v) => Math.abs(v) >= 10000 ? (v / 10000).toFixed(1) + '万' : Math.round(v).toLocaleString()
    : k.fmt === 'pct'
      ? (v) => v.toFixed(1) + '%'
      : (v) => Math.round(v).toLocaleString();
  chart.setOption({
    grid: { left: 8, right: 8, top: 28, bottom: 22 },
    xAxis: { type: 'category', data: D.monthsShort,
      axisLine: { show: false }, axisTick: { show: false },
      axisLabel: { fontSize: 11, color: '#64748b', fontWeight: 500 }, boundaryGap: true },
    yAxis: { type: 'value', show: false, scale: true },
    tooltip: { trigger: 'axis', formatter: (params) => params[0].name + '：' + fmtFn(params[0].value) },
    series: [{ type: 'line', data: k.values, smooth: true, symbol: 'circle', symbolSize: 7,
      lineStyle: { color: k.color, width: 2.5 }, itemStyle: { color: k.color },
      areaStyle: { color: k.color, opacity: 0.12 },
      label: { show: true, position: 'top', fontSize: 12, color: k.color, fontWeight: 700, distance: 6, formatter: (p) => fmtFn(p.value) } }],
  });
  window.addEventListener('resize', () => chart.resize());
});

(function chart1() {
  const el = document.getElementById('cmp-chart1');
  if (!el) return;
  const chart = echarts.init(el);
  const totals = D.series.grandTotal;
  chart.setOption({
    tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' },
      valueFormatter: (v) => v == null ? '—' : (Math.abs(v) >= 10000 ? (v / 10000).toFixed(1) + ' 万' : v.toLocaleString()) },
    legend: { top: 0, icon: 'roundRect', itemWidth: 14, itemHeight: 8 },
    grid: { left: 56, right: 56, top: 56, bottom: 32 },
    xAxis: { type: 'category', data: D.monthsShort,
      axisLine: { lineStyle: { color: '#e2e8f0' } }, axisLabel: { fontSize: 13, fontWeight: 600 } },
    yAxis: [{ type: 'value', name: '金额（元）', nameTextStyle: { color: '#64748b' },
      axisLabel: { formatter: (v) => v >= 10000 ? (v / 10000) + '万' : v }, splitLine: { lineStyle: { color: '#f1f5f9' } } },
      { type: 'value', name: '总收入折线', nameTextStyle: { color: '#64748b' },
      axisLabel: { formatter: (v) => v >= 10000 ? (v / 10000).toFixed(0) + '万' : v }, splitLine: { show: false } }],
    series: [
      { name: '病房', type: 'bar', stack: 'rev', data: D.series.inpatientRev, itemStyle: { color: COLOR.inpatient }, barWidth: '45%' },
      { name: '门诊', type: 'bar', stack: 'rev', data: D.series.outpatientRev, itemStyle: { color: COLOR.outpatient } },
      { name: '会诊门诊', type: 'bar', stack: 'rev', data: D.series.consultORev, itemStyle: { color: COLOR.consultO } },
      { name: '会诊住院', type: 'bar', stack: 'rev', data: D.series.consultIRev, itemStyle: { color: COLOR.consultI },
        label: topPctLabel(totals), markLine: avgMarkLine(totals, '总收入均值') },
      { name: '总收入', type: 'line', yAxisIndex: 1, data: totals, smooth: true, symbol: 'circle', symbolSize: 8,
        lineStyle: { color: '#1e3a8a', width: 3 }, itemStyle: { color: '#1e3a8a' },
        label: { show: true, position: 'top', fontSize: 11, fontWeight: 600, color: '#1e3a8a',
          formatter: (p) => p.value == null ? '' : (p.value / 10000).toFixed(0) + '万' }, z: 10 },
    ],
  });
  window.addEventListener('resize', () => chart.resize());
})();

(function chart2() {
  const el = document.getElementById('cmp-chart2');
  if (!el) return;
  const chart = echarts.init(el);
  chart.setOption({
    tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' },
      valueFormatter: (v) => v == null ? '—' : (Math.abs(v) >= 10000 ? (v / 10000).toFixed(1) + ' 万' : v.toLocaleString() + (Math.abs(v) < 100 ? '%' : '')) },
    legend: { top: 0 },
    grid: { left: 56, right: 56, top: 56, bottom: 32 },
    xAxis: { type: 'category', data: D.monthsShort,
      axisLine: { lineStyle: { color: '#e2e8f0' } }, axisLabel: { fontSize: 13, fontWeight: 600 } },
    yAxis: [{ type: 'value', name: '金额（元）', nameTextStyle: { color: '#64748b' },
        axisLabel: { formatter: (v) => v >= 10000 ? (v / 10000) + '万' : v }, splitLine: { lineStyle: { color: '#f1f5f9' } } },
      { type: 'value', name: '通州占比 (%)', nameTextStyle: { color: COLOR.tongzhou },
        axisLabel: { formatter: '{value}%', color: COLOR.tongzhou }, splitLine: { show: false },
        min: 0, max: (value) => Math.ceil(value.max + 5) }],
    series: [
      { name: '西直门', type: 'bar', data: D.series.grandXizhimen, itemStyle: { color: COLOR.xizhimen }, barWidth: '32%',
        label: topPctLabel(D.series.grandXizhimen), markLine: avgMarkLine(D.series.grandXizhimen, '西直门均值') },
      { name: '通州', type: 'bar', data: D.series.grandTongzhou, itemStyle: { color: COLOR.tongzhou },
        label: topPctLabel(D.series.grandTongzhou) },
      { name: '通州占比', type: 'line', yAxisIndex: 1, data: D.series.tongzhouRatio.map(v => v == null ? null : +v.toFixed(2)),
        smooth: true, symbol: 'circle', symbolSize: 8,
        lineStyle: { color: COLOR.tongzhou, width: 3 }, itemStyle: { color: COLOR.tongzhou },
        label: { show: true, position: 'top', fontSize: 11, fontWeight: 600, color: COLOR.tongzhou,
          formatter: (p) => p.value == null ? '' : p.value + '%' }, z: 10 },
    ],
  });
  window.addEventListener('resize', () => chart.resize());
})();

(function chart3() {
  const el = document.getElementById('cmp-chart3');
  if (!el) return;
  const chart = echarts.init(el);
  chart.setOption({
    tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' } },
    legend: { top: 0 },
    grid: { left: 56, right: 56, top: 56, bottom: 32 },
    xAxis: { type: 'category', data: D.monthsShort,
      axisLine: { lineStyle: { color: '#e2e8f0' } }, axisLabel: { fontSize: 13, fontWeight: 600 } },
    yAxis: [{ type: 'value', name: '例数', splitLine: { lineStyle: { color: '#f1f5f9' } } },
      { type: 'value', name: '外检总例数', splitLine: { show: false } }],
    series: [
      { name: '外检总例数', type: 'bar', data: D.series.waijianCases, itemStyle: { color: COLOR.primary }, barWidth: '24%',
        label: topPctLabel(D.series.waijianCases), markLine: avgMarkLine(D.series.waijianCases, '外检均值') },
      { name: '通州外检例数', type: 'bar', data: D.series.tzWaijianCases, itemStyle: { color: COLOR.tongzhou } },
      { name: '冰冻例数', type: 'bar', data: D.series.frozenCases, itemStyle: { color: COLOR.freeze } },
      { name: '外检趋势', type: 'line', yAxisIndex: 1, data: D.series.waijianCases,
        smooth: true, symbol: 'circle', symbolSize: 8,
        lineStyle: { color: '#1e3a8a', width: 3 }, itemStyle: { color: '#1e3a8a' }, z: 10 },
    ],
  });
  window.addEventListener('resize', () => chart.resize());
})();

(function chart4() {
  const el = document.getElementById('cmp-chart4');
  if (!el) return;
  const chart = echarts.init(el);
  chart.setOption({
    tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' }, valueFormatter: (v) => v == null ? '—' : v.toLocaleString() },
    legend: { top: 0 },
    grid: { left: 56, right: 56, top: 56, bottom: 32 },
    xAxis: { type: 'category', data: D.monthsShort,
      axisLine: { lineStyle: { color: '#e2e8f0' } }, axisLabel: { fontSize: 13, fontWeight: 600 } },
    yAxis: [{ type: 'value', name: '会诊金额（元）',
        axisLabel: { formatter: (v) => v >= 10000 ? (v / 10000).toFixed(1) + '万' : v }, splitLine: { lineStyle: { color: '#f1f5f9' } } },
      { type: 'value', name: '会诊总人次', splitLine: { show: false } }],
    series: [
      { name: '门诊会诊金额', type: 'bar', stack: 'consult', data: D.series.consultOAmt, itemStyle: { color: COLOR.consultO }, barWidth: '40%' },
      { name: '住院会诊金额', type: 'bar', stack: 'consult', data: D.series.consultIAmt, itemStyle: { color: COLOR.consultI },
        label: topPctLabel(D.series.consultTotalAmt), markLine: avgMarkLine(D.series.consultTotalAmt, '会诊均值') },
      { name: '会诊总人次', type: 'line', yAxisIndex: 1, data: D.series.consultTotalQty,
        smooth: true, symbol: 'circle', symbolSize: 8,
        lineStyle: { color: '#1e3a8a', width: 3 }, itemStyle: { color: '#1e3a8a' },
        label: { show: true, position: 'top', fontSize: 11, fontWeight: 600, color: '#1e3a8a',
          formatter: (p) => p.value == null ? '' : p.value.toLocaleString() + ' 人次' }, z: 10 },
    ],
  });
  window.addEventListener('resize', () => chart.resize());
})();

function buildSmallChart(elId, name, data, color) {
  const el = document.getElementById(elId);
  if (!el) return;
  const chart = echarts.init(el);
  chart.setOption({
    tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' },
      valueFormatter: (v) => v == null ? '—' : v.toLocaleString() + ' 例' },
    grid: { left: 50, right: 16, top: 36, bottom: 28 },
    xAxis: { type: 'category', data: D.monthsShort,
      axisLine: { lineStyle: { color: '#e2e8f0' } },
      axisLabel: { fontSize: 12, fontWeight: 600 } },
    yAxis: { type: 'value', name: '例数', nameTextStyle: { color: '#64748b', fontSize: 11 },
      splitLine: { lineStyle: { color: '#f1f5f9' } } },
    series: [
      { name, type: 'bar', data, itemStyle: { color, borderRadius: [4, 4, 0, 0] }, barWidth: '50%',
        label: topPctLabel(data),
        markLine: avgMarkLine(data, '均值') },
      { name: name + ' 趋势', type: 'line', data,
        smooth: true, symbol: 'none',
        lineStyle: { color: '#1e3a8a', width: 2, opacity: 0.5 },
        z: 1 },
    ],
  });
  window.addEventListener('resize', () => chart.resize());
}

(function chart5() {
  buildSmallChart('cmp-chart5a', '妇科细胞审核', D.series.fukeReview, COLOR.fuke);
  buildSmallChart('cmp-chart5b', '肾内切片审核', D.series.shenneiReview, COLOR.shennei);
  buildSmallChart('cmp-chart5c', '基因检测', D.series.genetic, COLOR.genetic);
})();

}

if (typeof echarts === 'undefined') {
  var s = document.createElement('script');
  s.src = 'https://cdn.staticfile.org/echarts/5.4.3/echarts.min.js';
  s.onload = initCompareCharts;
  document.head.appendChild(s);
} else {
  initCompareCharts();
}
})();
<\/script>
`;
}

function generateCompareReport(monthly) {
  monthly.sort((a, b) => a.meta.sortKey - b.meta.sortKey);
  const period = `${monthly[0].meta.period} ~ ${monthly[monthly.length - 1].meta.period}`;
  return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<title>病理科月度工作量对比 · ${period}</title>
<style>${compareCSS()}</style>
</head>
<body style="margin:0;background:#f8fafc;">
${compareBodyInner(monthly)}
${compareScripts(monthly)}
</body>
</html>`;
}

function generateCompareReportBody(monthly) {
  monthly.sort((a, b) => a.meta.sortKey - b.meta.sortKey);
  return `<style>${compareCSS()}</style>${compareBodyInner(monthly)}${compareScripts(monthly)}`;
}

module.exports = { generateCompareReport, generateCompareReportBody };
