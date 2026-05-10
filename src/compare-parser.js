const ExcelJS = require('exceljs');

function cellStr(v) {
  if (v == null) return '';
  if (typeof v === 'object') {
    if (v.result != null) return String(v.result).trim();
    if (v.text != null) return String(v.text).trim();
    if (v.richText) return v.richText.map(r => r.text).join('').trim();
    if (v.formula) return '=' + v.formula;
    return '';
  }
  return String(v).trim();
}

function pickNum(text) {
  if (text == null) return null;
  const s = String(text).replace(/,/g, '').trim();
  if (!s) return null;
  const eq = s.match(/=\s*([\d.]+)\s*$/);
  if (eq) return parseFloat(eq[1]);
  const m = s.match(/([\d.]+)/);
  return m ? parseFloat(m[1]) : null;
}

function parseMult(text) {
  const s = String(text || '').trim();
  const m = s.match(/(\d+)\s*\*\s*([\d.]+)\s*=\s*([\d.]+)/);
  if (m) return { price: parseInt(m[1]), qty: parseFloat(m[2]), amount: parseFloat(m[3]) };
  return null;
}

function parsePrefixedAmount(text) {
  if (!text) return null;
  const s = String(text).trim();
  const m = s.match(/(?:总|通|西|合计)?\s*([\d,.]+)/);
  return m ? parseFloat(m[1].replace(/,/g, '')) : null;
}

function extractPeriod(s) {
  const m = String(s || '').match(/(\d{4})\s*年\s*(\d{1,2})\s*月/);
  if (!m) return { period: '', year: null, month: null, sortKey: 0 };
  const year = parseInt(m[1]);
  const month = parseInt(m[2]);
  return {
    period: `${year}年${month}月`,
    year,
    month,
    sortKey: year * 100 + month,
  };
}

function parseWorkloadCount(wb) {
  const ws = wb.getWorksheet('工作量-数量') || wb.getWorksheet(1);
  if (!ws) return null;

  const result = {
    waijian_cases: null, waijian_blocks: null,
    frozen_cases: null, frozen_blocks: null,
    tongzhou_waijian_cases: null, tongzhou_waijian_blocks: null,
    tongzhou_frozen_cases: null, tongzhou_frozen_blocks: null,
    fuke_review: null, shennei_review: null,
    molecular: null, genetic: null,
  };

  ws.eachRow((row, rowNum) => {
    if (rowNum <= 2) return;
    const name = cellStr(row.getCell(1).value);
    const v2 = pickNum(cellStr(row.getCell(2).value));
    const v3 = pickNum(cellStr(row.getCell(3).value));
    const v4 = pickNum(cellStr(row.getCell(4).value));
    const v5 = pickNum(cellStr(row.getCell(5).value));

    if (name.includes('外检总量')) {
      result.waijian_cases = v2; result.waijian_blocks = v3;
      result.frozen_cases = v4; result.frozen_blocks = v5;
    } else if (name.includes('通州外检')) {
      result.tongzhou_waijian_cases = v2; result.tongzhou_waijian_blocks = v3;
      result.tongzhou_frozen_cases = v4; result.tongzhou_frozen_blocks = v5;
    } else if (name.includes('妇科细胞审核')) {
      result.fuke_review = v2;
    } else if (name.includes('肾内切片审核')) {
      result.shennei_review = v2;
    } else if (name.includes('分子病理')) {
      result.molecular = v2;
    } else if (name.includes('基因检测')) {
      result.genetic = v2;
    }
  });

  return result;
}

function parseConsultation(wb) {
  const ws = wb.getWorksheet('会诊统计') || wb.getWorksheet(2);
  if (!ws) return null;

  const sec = {
    outpatient: { high_xz: null, normal_xz: null, high_tz: null, normal_tz: null, subtotal: null },
    inpatient: { high_xz: null, normal_xz: null, high_tz: null, normal_tz: null, subtotal: null },
    grand_xz: null, grand_tz: null, grand_total: null,
  };

  let section = '';
  let rowInSec = 0;

  ws.eachRow((row, rowNum) => {
    if (rowNum <= 2) return;
    const c1 = cellStr(row.getCell(1).value);
    const c2 = cellStr(row.getCell(2).value);
    const c3 = cellStr(row.getCell(3).value);
    const c4 = cellStr(row.getCell(4).value);

    if (c1.includes('门诊会诊')) { section = 'outpatient'; rowInSec = 0; }
    else if (c1.includes('住院会诊')) { section = 'inpatient'; rowInSec = 0; }
    else if (c1.includes('共计') || c1.includes('合计')) {
      sec.grand_xz = parsePrefixedAmount(c2);
      sec.grand_tz = parsePrefixedAmount(c3);
      sec.grand_total = parsePrefixedAmount(c4);
      return;
    }

    const target = section === 'outpatient' ? sec.outpatient : section === 'inpatient' ? sec.inpatient : null;
    if (!target) return;

    const xz = parseMult(c2);
    const tz = parseMult(c3);
    const sub = parsePrefixedAmount(c4);

    if (rowInSec === 0) {
      if (xz) target.high_xz = xz;
      if (tz) target.high_tz = tz;
    } else {
      if (xz) target.normal_xz = xz;
      if (tz) target.normal_tz = tz;
      if (sub != null) target.subtotal = sub;
    }
    rowInSec++;
  });

  function qtySum(s) {
    return (s.high_xz?.qty || 0) + (s.normal_xz?.qty || 0)
      + (s.high_tz?.qty || 0) + (s.normal_tz?.qty || 0);
  }
  function amtSum(s) {
    return (s.high_xz?.amount || 0) + (s.normal_xz?.amount || 0)
      + (s.high_tz?.amount || 0) + (s.normal_tz?.amount || 0);
  }

  return {
    outpatient_qty: qtySum(sec.outpatient),
    outpatient_amount: sec.outpatient.subtotal != null ? sec.outpatient.subtotal : amtSum(sec.outpatient),
    inpatient_qty: qtySum(sec.inpatient),
    inpatient_amount: sec.inpatient.subtotal != null ? sec.inpatient.subtotal : amtSum(sec.inpatient),
    consult_qty_total: qtySum(sec.outpatient) + qtySum(sec.inpatient),
    consult_amount_total: sec.grand_total != null
      ? sec.grand_total
      : (sec.outpatient.subtotal || amtSum(sec.outpatient)) + (sec.inpatient.subtotal || amtSum(sec.inpatient)),
    consult_xz: sec.grand_xz,
    consult_tz: sec.grand_tz,
  };
}

function parseRevenue(wb) {
  const ws = wb.getWorksheet('工作量-金额') || wb.getWorksheet(3);
  if (!ws) return null;

  const result = {
    inpatient_total: null, inpatient_tongzhou: null,
    outpatient_total: null, outpatient_tongzhou: null,
    consult_outpatient_total: null, consult_outpatient_tongzhou: null,
    consult_inpatient_total: null, consult_inpatient_tongzhou: null,
    grand_total: null, grand_tongzhou: null,
  };

  let section = '';

  ws.eachRow((row, rowNum) => {
    if (rowNum <= 3) return;
    const c1 = cellStr(row.getCell(1).value);
    const c2 = cellStr(row.getCell(2).value);
    const c3 = cellStr(row.getCell(3).value);
    const c4 = cellStr(row.getCell(4).value);

    if (c1 === '病房' && c3) {
      result.inpatient_total = parsePrefixedAmount(c3);
      result.inpatient_tongzhou = parsePrefixedAmount(c4);
      section = '';
      return;
    }
    if (c1 === '门诊' && c3) {
      result.outpatient_total = parsePrefixedAmount(c3);
      result.outpatient_tongzhou = parsePrefixedAmount(c4);
      section = '';
      return;
    }
    if (c1.includes('病理会诊') && c1.includes('门诊')) {
      section = 'co';
      const tot = parseMult(c2);
      const xz = parseMult(c3);
      const tz = parseMult(c4);
      result.consult_outpatient_total = tot ? tot.amount : null;
      const xzAmt = xz ? xz.amount : null;
      const tzAmt = tz ? tz.amount : null;
      result._co_high = { tot: tot?.amount || 0, xz: xzAmt || 0, tz: tzAmt || 0 };
      return;
    }
    if (c1.includes('病理会诊') && c1.includes('住院')) {
      section = 'ci';
      const tot = parseMult(c2);
      const xz = parseMult(c3);
      const tz = parseMult(c4);
      result.consult_inpatient_total = tot ? tot.amount : null;
      result._ci_high = { tot: tot?.amount || 0, xz: xz?.amount || 0, tz: tz?.amount || 0 };
      return;
    }
    if (c1 === '合计' && c3) {
      result.grand_total = parsePrefixedAmount(c3);
      result.grand_tongzhou = parsePrefixedAmount(c4);
      section = '';
      return;
    }

    if (section === 'co' || section === 'ci') {
      const tot = parseMult(c2);
      const xz = parseMult(c3);
      const tz = parseMult(c4);
      const high = section === 'co' ? result._co_high : result._ci_high;
      const totSum = (tot?.amount || 0) + (high?.tot || 0);
      const tzSum = (tz?.amount || 0) + (high?.tz || 0);
      if (section === 'co') {
        result.consult_outpatient_total = totSum;
        result.consult_outpatient_tongzhou = tzSum;
      } else {
        result.consult_inpatient_total = totSum;
        result.consult_inpatient_tongzhou = tzSum;
      }
      section = '';
    }
  });

  delete result._co_high;
  delete result._ci_high;
  return result;
}

async function parseCompareExcel(filePath) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);

  const titleCell = cellStr(wb.worksheets[0]?.getCell('A1').value);
  const meta = extractPeriod(titleCell);

  return {
    meta,
    workload: parseWorkloadCount(wb),
    consultation: parseConsultation(wb),
    revenue: parseRevenue(wb),
  };
}

module.exports = { parseCompareExcel };
