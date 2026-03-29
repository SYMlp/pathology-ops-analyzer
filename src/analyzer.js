function analyze(parsedData) {
  const { meta, indicators, rankings, procurement, dispatch } = parsedData;
  const cm = meta.month;
  const pm = cm > 1 ? cm - 1 : null;

  return {
    meta,
    summary:     buildSummary(indicators, rankings, cm, pm),
    revenue:     buildRevenueAnalysis(indicators, cm, pm),
    expenses:    buildExpenseAnalysis(indicators, cm, pm),
    efficiency:  buildEfficiencyAnalysis(indicators, rankings, cm, pm),
    workload:    buildWorkloadAnalysis(indicators, cm, pm),
    rankings:    { ...rankings },
    procurement: buildProcurementAnalysis(procurement),
    dispatch:    buildDispatchAnalysis(dispatch),
    insights:    [],
  };
}

function postAnalyze(result) {
  result.insights = generateInsights(result);
  return result;
}

function v(indicators, key, month) {
  return indicators[key]?.monthly[month] ?? null;
}
function prevAvg(indicators, key) {
  return indicators[key]?.prev_year_avg ?? null;
}
function pct(curr, prev) {
  if (prev === null || prev === 0 || curr === null) return null;
  return (curr - prev) / Math.abs(prev);
}
function sum(...vals) {
  return vals.reduce((s, v) => s + (v || 0), 0);
}
function safeDiv(a, b) {
  if (b === null || b === 0 || a === null) return null;
  return a / b;
}

function buildSummary(indicators, rankings, cm, pm) {
  const staffCount    = v(indicators, 'staff_count', cm);
  const totalIncome   = v(indicators, 'operation_income_total', cm);
  const totalExpense  = v(indicators, 'total_expense', cm);
  const prevIncome    = pm ? v(indicators, 'operation_income_total', pm) : null;
  const prevExpense   = pm ? v(indicators, 'total_expense', pm) : null;
  const prevYearIncome  = prevAvg(indicators, 'operation_income_total');
  const prevYearExpense = prevAvg(indicators, 'total_expense');

  const workloadKeys = ['expert_outpatient', 'special_outpatient', 'general_outpatient'];
  const totalOutpatient = sum(...workloadKeys.map(k => v(indicators, k, cm)));
  const prevOutpatient  = pm ? sum(...workloadKeys.map(k => v(indicators, k, pm))) : null;
  const prevYearOutpatient = sum(...workloadKeys.map(k => prevAvg(indicators, k) || 0));

  const operationSurplus = (totalIncome && totalExpense) ? totalIncome - totalExpense : null;
  const prevSurplus = (prevIncome && prevExpense) ? prevIncome - prevExpense : null;

  return {
    staff_count: staffCount,
    total_income: totalIncome,
    total_expense: totalExpense,
    operation_surplus: operationSurplus,
    income_mom_change: pct(totalIncome, prevIncome),
    expense_mom_change: pct(totalExpense, prevExpense),
    surplus_mom_change: pct(operationSurplus, prevSurplus),
    income_yoy_change: pct(totalIncome, prevYearIncome),
    expense_yoy_change: pct(totalExpense, prevYearExpense),
    cost_revenue_ratio: safeDiv(totalExpense, totalIncome),
    prev_cost_revenue_ratio: safeDiv(prevExpense, prevIncome),
    hospital_cost_revenue_ratio: rankings.cost_revenue_ratio ?? null,
    total_outpatient: totalOutpatient,
    outpatient_mom_change: pct(totalOutpatient, prevOutpatient),
    outpatient_yoy_change: pct(totalOutpatient, prevYearOutpatient),
    per_capita_outpatient: safeDiv(totalOutpatient, staffCount),
    per_capita_income: safeDiv(totalIncome, staffCount),
    satisfaction_score: rankings.satisfaction_score ?? null,
    satisfaction_rank: rankings.satisfaction_rank ?? null,
    change_rank: rankings.change_rank ?? null,
  };
}

const REVENUE_KEYS = [
  'outpatient_award_income', 'medical_service_fee', 'operation_income_total',
  'flow_income_total', 'material_fee', 'drug_fee',
];

function buildRevenueAnalysis(indicators, cm, pm) {
  const items = REVENUE_KEYS.map(key => {
    const ind = indicators[key];
    if (!ind) return null;
    const current = ind.monthly[cm] ?? null;
    const prev = pm ? (ind.monthly[pm] ?? null) : null;
    return {
      key, label: ind.label,
      prev_year_avg: ind.prev_year_avg,
      current_year_avg: ind.current_year_avg,
      monthly: ind.monthly,
      current,
      mom_change: pct(current, prev),
      yoy_change: pct(current, ind.prev_year_avg),
    };
  }).filter(Boolean);

  const trend = indicators.operation_income_total?.monthly || {};
  return { items, trend };
}

const CONSUMABLE_KEYS = [
  'non_billable_consumables', 'office_consumables', 'reagent_usage',
  'medical_waste', 'oxygen', 'equipment_consumables',
];
const OTHER_EXPENSE_KEYS = [
  'third_party_fee', 'personnel_salary', 'part_time_professor_cost',
  'equipment_depreciation', 'equipment_repair', 'basic_bonus', 'night_shift_fee',
];

function buildExpenseAnalysis(indicators, cm, pm) {
  const allKeys = [...CONSUMABLE_KEYS, ...OTHER_EXPENSE_KEYS];
  const items = allKeys.map(key => {
    const ind = indicators[key];
    if (!ind) return null;
    const current = ind.monthly[cm] ?? null;
    const prev = pm ? (ind.monthly[pm] ?? null) : null;
    return {
      key, label: ind.label,
      prev_year_avg: ind.prev_year_avg,
      current_year_avg: ind.current_year_avg,
      monthly: ind.monthly,
      current,
      mom_change: pct(current, prev),
      yoy_change: pct(current, ind.prev_year_avg),
      isConsumable: CONSUMABLE_KEYS.includes(key),
    };
  }).filter(Boolean);

  const groups = {
    '各类消耗': CONSUMABLE_KEYS,
    '人员成本': ['personnel_salary', 'basic_bonus', 'night_shift_fee', 'part_time_professor_cost'],
    '设备支出': ['equipment_depreciation', 'equipment_repair'],
    '第三方费用': ['third_party_fee'],
  };

  const composition = {};
  for (const [group, keys] of Object.entries(groups)) {
    const total = sum(...keys.map(k => v(indicators, k, cm)));
    if (total > 0) composition[group] = total;
  }

  const trend = indicators.total_expense?.monthly || {};
  return { items, composition, trend };
}

function buildEfficiencyAnalysis(indicators, rankings, cm, pm) {
  const totalIncome   = v(indicators, 'operation_income_total', cm);
  const totalExpense  = v(indicators, 'total_expense', cm);
  const staffCount    = v(indicators, 'staff_count', cm);
  const salary        = v(indicators, 'personnel_salary', cm);
  const bonus         = v(indicators, 'basic_bonus', cm);
  const nightFee      = v(indicators, 'night_shift_fee', cm);
  const profCost      = v(indicators, 'part_time_professor_cost', cm);
  const consumableTotal = sum(...CONSUMABLE_KEYS.map(k => v(indicators, k, cm)));

  const personnelCost = sum(salary, bonus, nightFee, profCost);

  const totalAwardIncome = v(indicators, 'total_award_income', cm);

  return {
    cost_revenue_ratio: safeDiv(totalExpense, totalIncome),
    award_cost_ratio: safeDiv(totalExpense, totalAwardIncome),
    personnel_cost_ratio: safeDiv(personnelCost, totalIncome),
    consumable_per_100_income: totalIncome ? (consumableTotal / totalIncome) * 100 : null,
    per_capita_income: safeDiv(totalIncome, staffCount),
    per_capita_surplus: safeDiv(totalIncome - totalExpense, staffCount),
    operation_margin: safeDiv(totalIncome - totalExpense, totalIncome),
    satisfaction_score: rankings.satisfaction_score ?? null,
    satisfaction_rank: rankings.satisfaction_rank ?? null,
    hospital_cost_revenue_ratio: rankings.cost_revenue_ratio ?? null,
  };
}

function buildWorkloadAnalysis(indicators, cm, pm) {
  const keys = ['expert_outpatient', 'special_outpatient', 'general_outpatient'];
  const items = keys.map(key => {
    const ind = indicators[key];
    if (!ind) return null;
    const current = ind.monthly[cm] ?? null;
    const prev = pm ? (ind.monthly[pm] ?? null) : null;
    return {
      key, label: ind.label,
      prev_year_avg: ind.prev_year_avg,
      current_year_avg: ind.current_year_avg,
      monthly: ind.monthly,
      current,
      mom_change: pct(current, prev),
    };
  }).filter(Boolean);

  const trend = {};
  for (let m = 1; m <= 12; m++) {
    const total = sum(...keys.map(k => v(indicators, k, m)));
    const hasData = keys.some(k => v(indicators, k, m) !== null);
    if (hasData) trend[m] = total;
  }

  const staffCount = v(indicators, 'staff_count', cm);
  const currentTotal = trend[cm] || 0;
  const prevTotal = pm && trend[pm] != null ? trend[pm] : null;

  return {
    items, trend,
    per_capita: safeDiv(currentTotal, staffCount),
    total: currentTotal,
    mom_change: pct(currentTotal, prevTotal),
  };
}

function buildProcurementAnalysis(procurement) {
  if (!procurement?.length) {
    return { total_amount: 0, total_count: 0, top_items: [], supplier_distribution: {} };
  }
  const totalAmount = procurement.reduce((s, r) => s + (r.amount || r.subtotal || 0), 0);
  const byItem = {};
  procurement.forEach(r => {
    const name = r.item_name || '未知';
    if (!byItem[name]) byItem[name] = { name, amount: 0, qty: 0 };
    byItem[name].amount += (r.amount || r.subtotal || 0);
    byItem[name].qty    += (r.qty || 0);
  });
  const bySupplier = {};
  procurement.forEach(r => {
    const s = r.supplier || '未知';
    bySupplier[s] = (bySupplier[s] || 0) + (r.amount || r.subtotal || 0);
  });
  return {
    total_amount: totalAmount,
    total_count: procurement.length,
    top_items: Object.values(byItem).sort((a, b) => b.amount - a.amount).slice(0, 10),
    supplier_distribution: bySupplier,
  };
}

function buildDispatchAnalysis(dispatch) {
  if (!dispatch?.length) {
    return { total_amount: 0, total_count: 0, top_items: [] };
  }
  const totalAmount = dispatch.reduce((s, r) => s + (r.amount || 0), 0);
  const byItem = {};
  dispatch.forEach(r => {
    const name = r.item_name || '未知';
    if (!byItem[name]) byItem[name] = { name, amount: 0, qty: 0 };
    byItem[name].amount += (r.amount || 0);
    byItem[name].qty    += (r.qty || 0);
  });
  return {
    total_amount: totalAmount,
    total_count: dispatch.length,
    top_items: Object.values(byItem).sort((a, b) => b.amount - a.amount).slice(0, 10),
  };
}

function generateInsights(result) {
  const ins = [];
  const { summary: s, expenses: e, rankings: r, efficiency: ef } = result;

  if (s.cost_revenue_ratio !== null) {
    const ratio = s.cost_revenue_ratio;
    if (ratio > 0.85)
      ins.push({ type: 'danger', category: 'risk', title: '支收比严重超标',
        message: `当月支收比 ${fmtPct(ratio)}，大幅超过 80% 警戒线。需立即启动成本控制专项，重点排查消耗品和人力成本。`, metric: fmtPct(ratio), benchmark: '≤70%' });
    else if (ratio > 0.8)
      ins.push({ type: 'danger', category: 'risk', title: '支收比超警戒线',
        message: `当月支收比 ${fmtPct(ratio)}，超过 80% 警戒线。建议分析各支出类别环比变化，寻找可控降幅空间。`, metric: fmtPct(ratio), benchmark: '≤70%' });
    else if (ratio > 0.7)
      ins.push({ type: 'warning', category: 'risk', title: '支收比接近关注线',
        message: `当月支收比 ${fmtPct(ratio)}，接近 70% 关注线，需保持关注。`, metric: fmtPct(ratio), benchmark: '≤70%' });
    else
      ins.push({ type: 'positive', category: 'highlight', title: '支收比健康',
        message: `当月支收比 ${fmtPct(ratio)}，处于健康区间，成本控制良好。`, metric: fmtPct(ratio), benchmark: '≤70%' });
  }

  if (ef.personnel_cost_ratio !== null) {
    const ratio = ef.personnel_cost_ratio;
    if (ratio > 0.40)
      ins.push({ type: 'warning', category: 'risk', title: '人力成本占比偏高',
        message: `人力成本占收入比 ${fmtPct(ratio)}，超过 35% 行业基准。需评估人员配置效率。`, metric: fmtPct(ratio), benchmark: '≤35%' });
    else if (ratio <= 0.30)
      ins.push({ type: 'positive', category: 'highlight', title: '人力成本控制优秀',
        message: `人力成本占收入比 ${fmtPct(ratio)}，低于 30%，人效表现优异。`, metric: fmtPct(ratio), benchmark: '≤35%' });
  }

  if (ef.consumable_per_100_income !== null) {
    const val = ef.consumable_per_100_income;
    if (val > 30)
      ins.push({ type: 'warning', category: 'risk', title: '百元收入消耗偏高',
        message: `百元医疗收入消耗卫生材料 ${val.toFixed(1)} 元，偏高。建议审查试剂领用和设备耗材的使用效率。`, metric: val.toFixed(1) + '元', benchmark: '≤25元' });
  }

  if (s.income_mom_change !== null) {
    const c = s.income_mom_change;
    if (c < -0.1)
      ins.push({ type: 'danger', category: 'risk', title: '收入大幅下滑',
        message: `收入环比下降 ${fmtPct(Math.abs(c))}，降幅较大。需分析门诊量变化、收费项目调整等原因。`, metric: fmtPct(Math.abs(c)), benchmark: '' });
    else if (c < -0.03)
      ins.push({ type: 'warning', category: 'risk', title: '收入小幅下滑',
        message: `收入环比下降 ${fmtPct(Math.abs(c))}，需关注趋势。`, metric: fmtPct(Math.abs(c)), benchmark: '' });
    else if (c > 0.05)
      ins.push({ type: 'positive', category: 'highlight', title: '收入稳步增长',
        message: `收入环比增长 ${fmtPct(c)}，运营态势良好。`, metric: '+' + fmtPct(c), benchmark: '' });
  }

  if (s.outpatient_mom_change !== null) {
    const c = s.outpatient_mom_change;
    if (c < -0.2)
      ins.push({ type: 'danger', category: 'risk', title: '门诊量显著下降',
        message: `门诊量环比下降 ${fmtPct(Math.abs(c))}。需排查是否受节假日、季节因素影响，或存在患者分流问题。`, metric: fmtPct(Math.abs(c)), benchmark: '' });
    else if (c < -0.05)
      ins.push({ type: 'warning', category: 'risk', title: '门诊量有所下降',
        message: `门诊量环比下降 ${fmtPct(Math.abs(c))}。`, metric: fmtPct(Math.abs(c)), benchmark: '' });
    else if (c > 0.1)
      ins.push({ type: 'positive', category: 'highlight', title: '门诊量增长明显',
        message: `门诊量环比增长 ${fmtPct(c)}，工作量上升。需关注人均负荷是否合理。`, metric: '+' + fmtPct(c), benchmark: '' });
  }

  if (s.operation_surplus !== null) {
    if (s.operation_surplus > 0)
      ins.push({ type: 'positive', category: 'highlight', title: '实现正结余',
        message: `当月运营结余 ${fmtMoney(s.operation_surplus)}，科室运营健康。`, metric: fmtMoney(s.operation_surplus), benchmark: '' });
    else
      ins.push({ type: 'danger', category: 'risk', title: '运营亏损',
        message: `当月运营结余为负（${fmtMoney(s.operation_surplus)}），需紧急关注。`, metric: fmtMoney(s.operation_surplus), benchmark: '' });
  }

  if (r.satisfaction_rank) {
    if (r.satisfaction_rank <= 5)
      ins.push({ type: 'positive', category: 'highlight', title: '满意度领先',
        message: `全院满意度排名第 ${r.satisfaction_rank} 名，处于前列。`, metric: `第${r.satisfaction_rank}名`, benchmark: '' });
    else if (r.satisfaction_rank > 15)
      ins.push({ type: 'warning', category: 'suggestion', title: '满意度需提升',
        message: `满意度排名第 ${r.satisfaction_rank} 名，建议分析患者反馈改进服务流程。`, metric: `第${r.satisfaction_rank}名`, benchmark: '' });
  }

  if (s.cost_revenue_ratio !== null && s.prev_cost_revenue_ratio !== null) {
    const diff = s.cost_revenue_ratio - s.prev_cost_revenue_ratio;
    if (diff > 0.05)
      ins.push({ type: 'warning', category: 'suggestion', title: '支收比恶化趋势',
        message: `支收比较上月上升 ${(diff * 100).toFixed(1)} 个百分点，建议逐项排查支出增长来源。`, metric: '+' + (diff * 100).toFixed(1) + 'pp', benchmark: '' });
    else if (diff < -0.03)
      ins.push({ type: 'positive', category: 'highlight', title: '支收比改善',
        message: `支收比较上月下降 ${(Math.abs(diff) * 100).toFixed(1)} 个百分点，成本控制有效。`, metric: (diff * 100).toFixed(1) + 'pp', benchmark: '' });
  }

  return ins;
}

function fmtPct(ratio) { return (ratio * 100).toFixed(1) + '%'; }
function fmtMoney(val) {
  if (val == null) return '--';
  if (Math.abs(val) >= 10000) return (val / 10000).toFixed(2) + '万';
  return val.toFixed(2) + '元';
}

module.exports = { analyze, postAnalyze };
