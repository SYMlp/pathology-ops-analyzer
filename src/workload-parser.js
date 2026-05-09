const ExcelJS = require('exceljs');

function cellNum(cell) {
  if (cell == null) return null;
  if (typeof cell === 'number') return cell;
  const s = String(cell).replace(/,/g, '').trim();
  if (!s) return null;

  const eqMatch = s.match(/=\s*([\d.]+)\s*$/);
  if (eqMatch) return parseFloat(eqMatch[1]);

  const numMatch = s.match(/([\d.]+)/);
  if (numMatch) return parseFloat(numMatch[1]);
  return null;
}

function cellStr(cell) {
  if (cell == null) return '';
  if (cell.result != null) return String(cell.result).trim();
  return String(cell).trim();
}

function parseMultExpr(text) {
  if (!text) return null;
  const s = String(text).trim();
  const m = s.match(/(\d+)\s*\*\s*([\d.]+)\s*=\s*([\d.]+)/);
  if (m) return { price: parseInt(m[1]), qty: parseFloat(m[2]), amount: parseFloat(m[3]) };

  const amtMatch = s.match(/^[\d.]+$/);
  if (amtMatch) return { price: 0, qty: 0, amount: parseFloat(amtMatch[0]) };
  return null;
}

function parsePrefixedAmount(text) {
  if (!text) return null;
  const s = String(text).trim();
  const m = s.match(/(?:总|通|西|合计)?\s*([\d,.]+)/);
  return m ? parseFloat(m[1].replace(/,/g, '')) : null;
}

async function parseWorkloadExcel(filePath) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);

  const workload = parseWorkloadSheet(wb);
  const revenue = parseRevenueSheet(wb);
  const consultation = parseConsultationSheet(wb);
  const period = workload?.period || revenue?.period || consultation?.period || '';

  return {
    meta: { period, department: '病理科' },
    workload,
    revenue,
    consultation,
  };
}

function parseWorkloadSheet(wb) {
  const ws = wb.getWorksheet('工作量统计') || wb.getWorksheet(1);
  if (!ws) return null;

  const titleCell = cellStr(ws.getCell('A1').value);
  const periodMatch = titleCell.match(/(\d{4})\s*年\s*(\d{1,2})\s*月/);
  const period = periodMatch ? `${periodMatch[1]}年${periodMatch[2]}月` : '';

  const items = [];
  const KNOWN_NAMES = ['外检总量', '通州外检', '妇科细胞审核', '肾内切片审核', '分子病理', '基因检测'];

  ws.eachRow((row, rowNum) => {
    if (rowNum <= 2) return;
    const name = cellStr(row.getCell(1).value);
    if (!name || name.startsWith('科主任') || name.startsWith('组长') || name.startsWith('报表人')) return;

    const isKnown = KNOWN_NAMES.some(k => name.includes(k));
    if (!isKnown && !cellNum(row.getCell(2).value)) return;

    items.push({
      name: name.trim(),
      cases: cellNum(row.getCell(2).value),
      blocks: cellNum(row.getCell(3).value),
      frozen_cases: cellNum(row.getCell(4).value),
      frozen_blocks: cellNum(row.getCell(5).value),
    });
  });

  return { period, items };
}

function parseRevenueSheet(wb) {
  const ws = wb.getWorksheet('收入金额') || wb.getWorksheet(2);
  if (!ws) return null;

  const titleCell = cellStr(ws.getCell('A1').value);
  const periodMatch = titleCell.match(/(\d{4})\s*年\s*(\d{1,2})\s*月/);
  const period = periodMatch ? `${periodMatch[1]}年${periodMatch[2]}月` : '';

  let inpatient = { total: null, tongzhou: null };
  let outpatient = { total: null, tongzhou: null };
  let consultOutpatient = { items: [], total: null, tongzhou: null };
  let consultInpatient = { items: [], total: null, tongzhou: null };
  const outpatientDetails = [];
  let grandTotal = { total: null, tongzhou: null };

  let section = '';
  let inDetailSection = false;

  ws.eachRow((row, rowNum) => {
    if (rowNum <= 3) return;
    const col1 = cellStr(row.getCell(1).value);

    if (col1.includes('病房') && !col1.includes('会诊')) {
      section = 'inpatient';
      inpatient.total = parsePrefixedAmount(cellStr(row.getCell(3).value));
      inpatient.tongzhou = parsePrefixedAmount(cellStr(row.getCell(4).value));
      return;
    }

    if (col1.match(/^门诊$/) || (col1.includes('门诊') && !col1.includes('会诊') && !col1.includes('具体') && !col1.includes('元档'))) {
      if (!col1.includes('具体') && !inDetailSection && !col1.includes('会诊')) {
        section = 'outpatient';
        outpatient.total = parsePrefixedAmount(cellStr(row.getCell(3).value));
        outpatient.tongzhou = parsePrefixedAmount(cellStr(row.getCell(4).value));
        return;
      }
    }

    if (col1.includes('会诊') && col1.includes('门诊')) {
      section = 'consult_outpatient';
      const totalExpr = parseMultExpr(cellStr(row.getCell(3).value));
      const tzExpr = parseMultExpr(cellStr(row.getCell(4).value));
      if (totalExpr) consultOutpatient.items.push({ ...totalExpr, source: 'total' });
      if (tzExpr) consultOutpatient.items.push({ ...tzExpr, source: 'tongzhou' });
      return;
    }
    if (col1.includes('会诊') && col1.includes('住院')) {
      section = 'consult_inpatient';
      const totalExpr = parseMultExpr(cellStr(row.getCell(3).value));
      const tzExpr = parseMultExpr(cellStr(row.getCell(4).value));
      if (totalExpr) consultInpatient.items.push({ ...totalExpr, source: 'total' });
      if (tzExpr) consultInpatient.items.push({ ...tzExpr, source: 'tongzhou' });
      return;
    }
    if (col1.includes('具体')) {
      section = 'detail';
      inDetailSection = true;
      return;
    }
    if (col1.includes('合计')) {
      grandTotal.total = parsePrefixedAmount(cellStr(row.getCell(3).value));
      grandTotal.tongzhou = parsePrefixedAmount(cellStr(row.getCell(4).value));
      return;
    }

    if (section === 'consult_outpatient' || section === 'consult_inpatient') {
      const totalExpr = parseMultExpr(cellStr(row.getCell(3).value));
      const tzExpr = parseMultExpr(cellStr(row.getCell(4).value));
      const target = section === 'consult_outpatient' ? consultOutpatient : consultInpatient;
      if (totalExpr) target.items.push({ ...totalExpr, source: 'total' });
      if (tzExpr) target.items.push({ ...tzExpr, source: 'tongzhou' });
    }

    if (inDetailSection) {
      const priceMatch = col1.match(/(\d+)/);
      if (priceMatch) {
        const price = parseInt(priceMatch[1]);
        const totalExpr = parseMultExpr(cellStr(row.getCell(3).value));
        const tzExpr = parseMultExpr(cellStr(row.getCell(4).value));
        outpatientDetails.push({
          price,
          total: totalExpr || { price, qty: cellNum(row.getCell(2).value), amount: cellNum(row.getCell(3).value) },
          tongzhou: tzExpr || { price, qty: 0, amount: cellNum(row.getCell(4).value) },
        });
      }
    }
  });

  consultOutpatient.total = consultOutpatient.items.filter(i => i.source === 'total').reduce((s, i) => s + (i.amount || 0), 0);
  consultOutpatient.tongzhou = consultOutpatient.items.filter(i => i.source === 'tongzhou').reduce((s, i) => s + (i.amount || 0), 0);
  consultInpatient.total = consultInpatient.items.filter(i => i.source === 'total').reduce((s, i) => s + (i.amount || 0), 0);
  consultInpatient.tongzhou = consultInpatient.items.filter(i => i.source === 'tongzhou').reduce((s, i) => s + (i.amount || 0), 0);

  return {
    period,
    inpatient,
    outpatient,
    consultOutpatient,
    consultInpatient,
    outpatientDetails,
    grandTotal,
  };
}

function parseConsultationSheet(wb) {
  const ws = wb.getWorksheet('会诊统计') || wb.getWorksheet(3);
  if (!ws) return null;

  const titleCell = cellStr(ws.getCell('A1').value);
  const periodMatch = titleCell.match(/(\d{4})\s*年\s*(\d{1,2})\s*月/);
  const period = periodMatch ? `${periodMatch[1]}年${periodMatch[2]}月` : '';

  const data = {
    outpatient: { xizhimen_high: null, xizhimen_normal: null, tongzhou_high: null, tongzhou_normal: null, subtotal: null },
    inpatient: { xizhimen_high: null, xizhimen_normal: null, tongzhou_high: null, tongzhou_normal: null, subtotal: null },
    total: { xizhimen: null, tongzhou: null, grand: null },
  };

  let section = '';
  let rowInSection = 0;

  ws.eachRow((row, rowNum) => {
    if (rowNum <= 2) return;
    const col1 = cellStr(row.getCell(1).value);
    const col2 = cellStr(row.getCell(2).value);
    const col3 = cellStr(row.getCell(3).value);
    const col4 = cellStr(row.getCell(4).value);

    if (col1.includes('门诊会诊')) {
      section = 'outpatient';
      rowInSection = 0;
    } else if (col1.includes('住院会诊')) {
      section = 'inpatient';
      rowInSection = 0;
    } else if (col1.includes('共计') || col1.includes('合计')) {
      section = 'total';
    }

    if (section === 'total') {
      data.total.xizhimen = parsePrefixedAmount(col2);
      data.total.tongzhou = parsePrefixedAmount(col3);
      data.total.grand = parsePrefixedAmount(col4);
      return;
    }

    const target = section === 'outpatient' ? data.outpatient : section === 'inpatient' ? data.inpatient : null;
    if (!target) return;

    const xzExpr = parseMultExpr(col2);
    const tzExpr = parseMultExpr(col3);
    const subtotal = parsePrefixedAmount(col4);

    if (rowInSection === 0) {
      if (xzExpr) target.xizhimen_high = xzExpr;
      if (tzExpr) target.tongzhou_high = tzExpr;
    } else {
      if (xzExpr) target.xizhimen_normal = xzExpr;
      if (tzExpr) target.tongzhou_normal = tzExpr;
      if (subtotal) target.subtotal = subtotal;
    }
    rowInSection++;
  });

  function sumSection(s) {
    return (s.xizhimen_high?.amount || 0) + (s.xizhimen_normal?.amount || 0)
      + (s.tongzhou_high?.amount || 0) + (s.tongzhou_normal?.amount || 0);
  }
  const opHasAny = data.outpatient.xizhimen_high || data.outpatient.xizhimen_normal
    || data.outpatient.tongzhou_high || data.outpatient.tongzhou_normal;
  const ipHasAny = data.inpatient.xizhimen_high || data.inpatient.xizhimen_normal
    || data.inpatient.tongzhou_high || data.inpatient.tongzhou_normal;
  if (!data.outpatient.subtotal && opHasAny) data.outpatient.subtotal = sumSection(data.outpatient);
  if (!data.inpatient.subtotal && ipHasAny) data.inpatient.subtotal = sumSection(data.inpatient);
  if (!data.total.grand) {
    data.total.grand = (data.outpatient.subtotal || 0) + (data.inpatient.subtotal || 0);
  }
  if (!data.total.xizhimen) {
    data.total.xizhimen =
      (data.outpatient.xizhimen_high?.amount || 0) + (data.outpatient.xizhimen_normal?.amount || 0) +
      (data.inpatient.xizhimen_high?.amount || 0) + (data.inpatient.xizhimen_normal?.amount || 0);
  }
  if (!data.total.tongzhou) {
    data.total.tongzhou =
      (data.outpatient.tongzhou_high?.amount || 0) + (data.outpatient.tongzhou_normal?.amount || 0) +
      (data.inpatient.tongzhou_high?.amount || 0) + (data.inpatient.tongzhou_normal?.amount || 0);
  }

  return { period, ...data };
}

module.exports = { parseWorkloadExcel };
