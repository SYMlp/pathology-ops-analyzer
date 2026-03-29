/**
 * 大屏 ECharts — 深色主题，依赖 window.echarts 与全局 initBigScreen(analysis)
 */
(function () {
  const ML = ['1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月'];
  const ACCENT = ['#38bdf8', '#34d399', '#fbbf24', '#a78bfa', '#22d3ee', '#fb7185', '#818cf8', '#2dd4bf'];
  const AXIS = '#64748b';
  const GRID = 'rgba(148,163,184,0.15)';
  const TXT = '#94a3b8';

  function getActiveMonths(trend) {
    return Object.keys(trend || {})
      .map(Number)
      .sort((a, b) => a - b);
  }

  function fmtV(v) {
    if (v >= 10000) return (v / 10000).toFixed(1) + '万';
    if (v >= 1000) return v.toLocaleString();
    return v;
  }

  window.initBigScreen = function (d) {
    if (!d || typeof echarts === 'undefined') return;

    const charts = [];
    function mount(id, option) {
      const el = document.getElementById(id);
      if (!el) return;
      const c = echarts.init(el);
      c.setOption(option);
      charts.push(c);
    }

    const { revenue, expenses, workload, procurement, efficiency, meta } = d;
    const rMonths = getActiveMonths(revenue.trend);
    const eMonths = getActiveMonths(expenses.trend);
    let allMonths = [...new Set([...rMonths, ...eMonths])].sort((a, b) => a - b);
    if (meta && meta.month >= 1) {
      allMonths = allMonths.filter(function (m) {
        return m <= meta.month;
      });
    }
    if (!allMonths.length && meta && meta.month >= 1) {
      allMonths = Array.from({ length: meta.month }, function (_, i) {
        return i + 1;
      });
    }

    if (allMonths.length) {
      mount('bs-chart-line', {
        backgroundColor: 'transparent',
        textStyle: { color: TXT },
        tooltip: {
          trigger: 'axis',
          backgroundColor: 'rgba(15,23,42,0.95)',
          borderColor: '#334155',
          textStyle: { color: '#f1f5f9' },
          formatter: function (p) {
            var h = p[0].axisValue;
            p.forEach(function (i) {
              h +=
                '<br/>' +
                i.marker +
                i.seriesName +
                ': ' +
                (i.value >= 10000 ? (i.value / 10000).toFixed(2) + '万' : i.value.toLocaleString() + '元');
            });
            return h;
          },
        },
        legend: { bottom: 0, textStyle: { color: TXT, fontSize: 10 } },
        grid: { top: 12, right: 8, bottom: 28, left: 44 },
        xAxis: {
          type: 'category',
          data: allMonths.map((m) => ML[m - 1]),
          axisLine: { lineStyle: { color: AXIS } },
          axisLabel: { color: TXT, fontSize: 10 },
          axisTick: { show: false },
        },
        yAxis: {
          type: 'value',
          scale: true,
          axisLabel: { color: TXT, fontSize: 10, formatter: fmtV },
          splitLine: { lineStyle: { color: GRID } },
          axisLine: { show: false },
        },
        series: [
          {
            name: '计奖收入',
            type: 'line',
            smooth: true,
            data: allMonths.map((m) => revenue.trend[m] ?? 0),
            lineStyle: { width: 2, color: '#34d399' },
            itemStyle: { color: '#34d399' },
            areaStyle: {
              color: { type: 'linear', x: 0, y: 0, x2: 0, y2: 1, colorStops: [
                { offset: 0, color: 'rgba(52,211,153,0.25)' },
                { offset: 1, color: 'rgba(52,211,153,0)' },
              ]},
            },
          },
          {
            name: '科室支出',
            type: 'line',
            smooth: true,
            data: allMonths.map((m) => expenses.trend[m] ?? 0),
            lineStyle: { width: 2, color: '#f87171' },
            itemStyle: { color: '#f87171' },
            areaStyle: {
              color: { type: 'linear', x: 0, y: 0, x2: 0, y2: 1, colorStops: [
                { offset: 0, color: 'rgba(248,113,113,0.2)' },
                { offset: 1, color: 'rgba(248,113,113,0)' },
              ]},
            },
          },
        ],
      });
    }

    if (Object.keys(expenses.composition || {}).length) {
      const pieData = Object.entries(expenses.composition).map(([name, value], i) => ({
        name,
        value,
        itemStyle: { color: ACCENT[i % ACCENT.length] },
      }));
      mount('bs-chart-pie', {
        backgroundColor: 'transparent',
        textStyle: { color: TXT },
        tooltip: {
          trigger: 'item',
          backgroundColor: 'rgba(15,23,42,0.95)',
          textStyle: { color: '#f1f5f9' },
          formatter: function (p) {
            return (
              p.marker +
              p.name +
              '<br/>' +
              (p.value >= 10000 ? (p.value / 10000).toFixed(2) + '万' : p.value.toLocaleString() + '元') +
              '<br/>' +
              p.percent.toFixed(1) +
              '%'
            );
          },
        },
        series: [
          {
            type: 'pie',
            radius: ['32%', '58%'],
            center: ['50%', '52%'],
            avoidLabelOverlap: true,
            label: { color: '#cbd5e1', fontSize: 9, formatter: '{b}\n{d}%' },
            labelLine: { lineStyle: { color: AXIS } },
            data: pieData,
          },
        ],
      });
    }

    let wMonths = getActiveMonths(workload.trend);
    if (meta && meta.month >= 1) {
      wMonths = wMonths.filter(function (m) {
        return m <= meta.month;
      });
    }
    if (wMonths.length) {
      const wData = wMonths.map((m) => ({
        value: workload.trend[m] ?? 0,
        itemStyle: {
          color: m === meta.month ? '#38bdf8' : 'rgba(56,189,248,0.35)',
          borderRadius: [3, 3, 0, 0],
        },
      }));
      mount('bs-chart-bar', {
        backgroundColor: 'transparent',
        textStyle: { color: TXT },
        tooltip: {
          trigger: 'axis',
          backgroundColor: 'rgba(15,23,42,0.95)',
          textStyle: { color: '#f1f5f9' },
        },
        grid: { top: 12, right: 8, bottom: 22, left: 36 },
        xAxis: {
          type: 'category',
          data: wMonths.map((m) => ML[m - 1]),
          axisLine: { lineStyle: { color: AXIS } },
          axisLabel: { color: TXT, fontSize: 10 },
          axisTick: { show: false },
        },
        yAxis: {
          type: 'value',
          axisLabel: { color: TXT, fontSize: 10 },
          splitLine: { lineStyle: { color: GRID } },
          axisLine: { show: false },
        },
        series: [{ type: 'bar', barWidth: '45%', data: wData, label: { show: true, position: 'top', fontSize: 9, color: TXT } }],
      });
    }

    if (efficiency.cost_revenue_ratio != null) {
      const val = +(efficiency.cost_revenue_ratio * 100).toFixed(1);
      const gaugeColor = val > 80 ? '#f87171' : val > 70 ? '#fbbf24' : '#34d399';
      mount('bs-chart-gauge', {
        backgroundColor: 'transparent',
        series: [
          {
            type: 'gauge',
            startAngle: 180,
            endAngle: 0,
            min: 0,
            max: 100,
            splitNumber: 5,
            center: ['50%', '70%'],
            radius: '95%',
            pointer: { show: true, length: '55%', width: 4, itemStyle: { color: gaugeColor } },
            progress: { show: true, width: 12, itemStyle: { color: gaugeColor } },
            axisLine: {
              lineStyle: {
                width: 12,
                color: [
                  [0.7, '#34d399'],
                  [0.8, '#fbbf24'],
                  [1, '#f87171'],
                ],
              },
            },
            axisTick: { show: false },
            splitLine: { length: 6, lineStyle: { width: 1, color: AXIS } },
            axisLabel: { distance: 14, fontSize: 9, color: TXT },
            detail: {
              valueAnimation: true,
              formatter: '{value}%',
              fontSize: 16,
              fontWeight: 'bold',
              color: gaugeColor,
              offsetCenter: [0, '0%'],
            },
            title: { show: true, offsetCenter: [0, '22%'], fontSize: 10, color: TXT },
            data: [{ value: val, name: '支收比' }],
          },
        ],
      });
    }

    window.addEventListener('resize', function () {
      charts.forEach(function (c) {
        c.resize();
      });
    });
  };
})();
