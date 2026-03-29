const ExcelJS = require('exceljs');

const INDICATOR_MAP = {
  '科室人数':         { key: 'staff_count',               category: 'basic' },
  '门诊计奖收入总计': { key: 'outpatient_award_income',    category: 'revenue' },
  '门诊医事服务费':   { key: 'medical_service_fee',        category: 'revenue' },
  '门诊运营收入总计': { key: 'operation_income_total',      category: 'revenue' },
  '门诊流水收入总计': { key: 'flow_income_total',           category: 'revenue' },
  '门诊材料费':       { key: 'material_fee',                category: 'revenue' },
  '门诊药费':         { key: 'drug_fee',                    category: 'revenue' },
  '非收费耗材':       { key: 'non_billable_consumables',    category: 'expense_consumable' },
  '办公耗材':         { key: 'office_consumables',          category: 'expense_consumable' },
  '试剂领用':         { key: 'reagent_usage',               category: 'expense_consumable' },
  '医疗垃圾':         { key: 'medical_waste',               category: 'expense_consumable' },
  '氧气':             { key: 'oxygen',                      category: 'expense_consumable' },
  '设备类耗材':       { key: 'equipment_consumables',       category: 'expense_consumable' },
  '第三方费用':       { key: 'third_party_fee',             category: 'expense_other' },
  '人员工资':         { key: 'personnel_salary',            category: 'expense_other' },
  '兼职教授成本':     { key: 'part_time_professor_cost',    category: 'expense_other' },
  '设备折旧':         { key: 'equipment_depreciation',      category: 'expense_other' },
  '设备维修费':       { key: 'equipment_repair',            category: 'expense_other' },
  '基本奖金':         { key: 'basic_bonus',                 category: 'expense_other' },
  '夜班费':           { key: 'night_shift_fee',             category: 'expense_other' },
  '科室计奖收入合计': { key: 'total_award_income',          category: 'summary' },
  '科室支出合计':     { key: 'total_expense',               category: 'summary' },
  '专家门诊量':       { key: 'expert_outpatient',           category: 'workload' },
  '特需门诊量':       { key: 'special_outpatient',          category: 'workload' },
  '普通门诊量':       { key: 'general_outpatient',          category: 'workload' },
};

const COL_PREV_YEAR_AVG = 2;
const COL_CURRENT_YEAR_AVG = 3;
const COL_MONTH_START = 4;

function parseNum(val) {
  if (val === null || val === undefined || val === '') return null;
  if (typeof val === 'object' && 'result' in val) val = val.result;
  if (val === null || val === undefined || val === '') return null;
  if (typeof val === 'number') return val;
  const cleaned = String(val).replace(/,/g, '').trim();
  const num = parseFloat(cleaned);
  return isNaN(num) ? null : num;
}

function stripPrefix(text) {
  return text
    .replace(/^[\d\.\s、（）()\-一二三四五六七八九十]+/, '')
    .replace(/（.*?）/, '')
    .trim();
}

function matchIndicator(cellText) {
  const cleaned = stripPrefix(cellText);
  for (const [name, info] of Object.entries(INDICATOR_MAP)) {
    if (cleaned === name || cellText.includes(name)) {
      return info;
    }
  }
  return null;
}

async function parseExcel(filePath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  const data = {
    meta: { department: '病理科', year: new Date().getFullYear(), month: 1 },
    indicators: {},
    rankings: {},
    procurement: [],
    dispatch: [],
  };

  const sheet1 = workbook.getWorksheet('运营指标') || workbook.getWorksheet(1);
  if (sheet1) parseOperationsSheet(sheet1, data);

  const sheet2 = workbook.getWorksheet('全院排名') || workbook.getWorksheet(2);
  if (sheet2) parseRankingsSheet(sheet2, data);

  const sheet3 = workbook.getWorksheet('采购入库') || workbook.getWorksheet(3);
  if (sheet3) parseTableSheet(sheet3, data.procurement, PROCUREMENT_COL_MAP);

  const sheet4 = workbook.getWorksheet('出库领用') || workbook.getWorksheet(4);
  if (sheet4) parseTableSheet(sheet4, data.dispatch, DISPATCH_COL_MAP);

  detectMeta(data, sheet1);
  return data;
}

function parseOperationsSheet(sheet, data) {
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber <= 2) return;

    const rawText = getCellText(row, 1);
    if (!rawText) return;

    const info = matchIndicator(rawText);
    if (!info) return;

    const monthly = {};
    let hasData = false;
    for (let m = 1; m <= 12; m++) {
      const val = parseNum(row.getCell(COL_MONTH_START + m - 1).value);
      if (val !== null) {
        monthly[m] = val;
        hasData = true;
      }
    }

    if (data.indicators[info.key] && !hasData && Object.keys(data.indicators[info.key].monthly).length > 0) {
      return; 
    }

    data.indicators[info.key] = {
      category: info.category,
      label: findLabelForKey(info.key),
      prev_year_avg: parseNum(row.getCell(COL_PREV_YEAR_AVG).value),
      current_year_avg: parseNum(row.getCell(COL_CURRENT_YEAR_AVG).value),
      monthly,
    };
  });
}

function findLabelForKey(key) {
  for (const [name, info] of Object.entries(INDICATOR_MAP)) {
    if (info.key === key) return name;
  }
  return key;
}

const RANKING_COL_MAP = {
  '5分满意度':   'satisfaction_score',
  '满意度排名':   'satisfaction_rank',
  '支收比':       'cost_revenue_ratio',
  '支收比增量':   'cost_revenue_ratio_change',
  '预约效率增量': 'appointment_efficiency_change',
  '增量和':       'total_change',
  '增量排序':     'change_rank',
  '科室字典':     'department',
};

function parseRankingsSheet(sheet, data) {
  const headerRow = sheet.getRow(1);
  const colMap = {};

  const sortedEntries = Object.entries(RANKING_COL_MAP)
    .sort((a, b) => b[0].length - a[0].length);

  headerRow.eachCell((cell, colNumber) => {
    const header = String(cell.value || '').trim();
    for (const [cn, en] of sortedEntries) {
      if (header === cn || (header.includes(cn) && !sortedEntries.some(
        ([longer]) => longer.length > cn.length && header.includes(longer)
      ))) {
        colMap[en] = colNumber;
        break;
      }
    }
  });

  const dataRow = sheet.getRow(2);
  for (const [key, col] of Object.entries(colMap)) {
    const val = dataRow.getCell(col).value;
    data.rankings[key] = (key === 'department') ? String(val || '') : parseNum(val);
  }
}

const PROCUREMENT_COL_MAP = {
  '订单号':     { field: 'order_id',   type: 'string' },
  '物料说明':   { field: 'item_name',  type: 'string' },
  '耗材名称':   { field: 'item_name',  type: 'string' },
  '购料说明':   { field: 'item_name',  type: 'string' },
  '规格':       { field: 'spec',       type: 'string' },
  '供应商':     { field: 'supplier',   type: 'string' },
  '入库数量':   { field: 'qty',        type: 'number' },
  '数量':       { field: 'qty',        type: 'number' },
  '单价':       { field: 'unit_price', type: 'number' },
  '入库金额':   { field: 'amount',     type: 'number' },
  '金额':       { field: 'amount',     type: 'number' },
  '小计金额':   { field: 'subtotal',   type: 'number' },
  '入库时间':   { field: 'date',       type: 'date' },
  '入库日期':   { field: 'date',       type: 'date' },
  '入期时间':   { field: 'date',       type: 'date' },
  '科室对应':   { field: 'department', type: 'string' },
  '科室':       { field: 'department', type: 'string' },
};

const DISPATCH_COL_MAP = {
  '序号':       { field: 'id',              type: 'string' },
  '物料说明':   { field: 'item_name',       type: 'string' },
  '物品名称':   { field: 'item_name',       type: 'string' },
  '费用名称':   { field: 'item_name',       type: 'string' },
  '名称':       { field: 'item_name',       type: 'string' },
  '规格':       { field: 'spec',            type: 'string' },
  '发出仓库':   { field: 'from_department', type: 'string' },
  '出库科室':   { field: 'to_department',   type: 'string' },
  '收费科室':   { field: 'to_department',   type: 'string' },
  '出库数量':   { field: 'qty',             type: 'number' },
  '数量':       { field: 'qty',             type: 'number' },
  '总成本':     { field: 'amount',          type: 'number' },
  '金额':       { field: 'amount',          type: 'number' },
  '费用':       { field: 'amount',          type: 'number' },
  '使用日期':   { field: 'date',            type: 'date' },
  '生产厂家':   { field: 'manufacturer',    type: 'string' },
  '科室对应':   { field: 'department',      type: 'string' },
  '科室':       { field: 'department',      type: 'string' },
};

function parseTableSheet(sheet, targetArray, colMapping) {
  const headerRow = sheet.getRow(1);
  const resolved = {};

  headerRow.eachCell((cell, colNumber) => {
    const header = String(cell.value || '').trim();
    for (const [cn, info] of Object.entries(colMapping)) {
      if (header.includes(cn) && !resolved[info.field]) {
        resolved[info.field] = { col: colNumber, type: info.type };
        break;
      }
    }
  });

  sheet.eachRow((row, rowNumber) => {
    if (rowNumber <= 1) return;
    const record = {};
    let hasContent = false;

    for (const [field, { col, type }] of Object.entries(resolved)) {
      const raw = row.getCell(col).value;
      if (type === 'date') {
        record[field] = raw instanceof Date ? raw : (raw ? new Date(raw) : null);
      } else if (type === 'number') {
        record[field] = parseNum(raw);
      } else {
        record[field] = raw ? String(raw).trim() : '';
      }
      if (raw !== null && raw !== undefined && raw !== '') hasContent = true;
    }

    if (hasContent) targetArray.push(record);
  });
}

function detectMeta(data, sheet) {
  let latestMonth = 1;
  for (const ind of Object.values(data.indicators)) {
    const months = Object.keys(ind.monthly).map(Number);
    if (months.length) latestMonth = Math.max(latestMonth, ...months);
  }
  data.meta.month = latestMonth;

  if (sheet) {
    const titleCell = String(sheet.getRow(1).getCell(1).value || '');
    const yearMatch = titleCell.match(/(\d{4})年/);
    if (yearMatch) data.meta.year = parseInt(yearMatch[1]);
    const deptMatch = titleCell.match(/(?:年\d+月|月)([\u4e00-\u9fa5]{2,}科)/);
    if (deptMatch) data.meta.department = deptMatch[1];
  }
}

function getCellText(row, colNumber) {
  const cell = colNumber === 0 ? row.getCell(1) : row.getCell(colNumber);
  const val = cell?.value;
  if (val === null || val === undefined) return '';
  if (typeof val === 'object' && val.richText) {
    return val.richText.map(r => r.text).join('').trim();
  }
  return String(val).trim();
}

module.exports = { parseExcel };
