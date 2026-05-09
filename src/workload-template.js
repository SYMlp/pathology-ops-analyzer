const ExcelJS = require('exceljs');
const path = require('path');

const BORDER_THIN = {
  top: { style: 'thin' }, bottom: { style: 'thin' },
  left: { style: 'thin' }, right: { style: 'thin' },
};
const FONT_HEADER = { bold: true, size: 11 };
const FONT_TITLE = { bold: true, size: 14 };
const HEADER_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
const SIGNATURE_FONT = { size: 11 };

function applyBorders(row, colCount) {
  for (let c = 1; c <= colCount; c++) {
    row.getCell(c).border = BORDER_THIN;
  }
}

async function generateWorkloadTemplate(outputPath) {
  const wb = new ExcelJS.Workbook();
  wb.creator = '病理科运营分析系统';

  buildWorkloadSheet(wb);
  buildRevenueSheet(wb);
  buildConsultationSheet(wb);
  buildInstructionsSheet(wb);

  const filePath = outputPath || path.join(__dirname, '..', 'templates', '月度工作报表模板.xlsx');
  await wb.xlsx.writeFile(filePath);
  return filePath;
}

function buildWorkloadSheet(wb) {
  const ws = wb.addWorksheet('工作量统计');
  const COL_COUNT = 5;

  ws.columns = [
    { key: 'name', width: 18 },
    { key: 'cases', width: 14 },
    { key: 'blocks', width: 14 },
    { key: 'frozen_cases', width: 14 },
    { key: 'frozen_blocks', width: 14 },
  ];

  ws.mergeCells('A1:E1');
  const titleCell = ws.getCell('A1');
  titleCell.value = '____年__月病理科工作量';
  titleCell.font = FONT_TITLE;
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getRow(1).height = 30;

  const headerRow = ws.getRow(2);
  headerRow.values = ['名称', '例数', '块数', '冰冻例数', '冰冻块数'];
  headerRow.font = FONT_HEADER;
  headerRow.eachCell(c => { c.fill = HEADER_FILL; c.border = BORDER_THIN; c.alignment = { horizontal: 'center' }; });

  const items = [
    { name: '外检总量' },
    { name: '通州外检' },
    { name: '妇科细胞审核' },
    { name: '肾内切片审核' },
    { name: '分子病理' },
    { name: '基因检测' },
  ];

  items.forEach((item, idx) => {
    const row = ws.getRow(idx + 3);
    row.getCell(1).value = item.name;
    applyBorders(row, COL_COUNT);
  });

  const sigRow = ws.getRow(items.length + 4);
  sigRow.getCell(1).value = '科主任：';
  sigRow.getCell(2).value = '组长：';
  sigRow.getCell(3).value = '报表人：';
  sigRow.getCell(4).value = '日期：';
  sigRow.font = SIGNATURE_FONT;
}

function buildRevenueSheet(wb) {
  const ws = wb.addWorksheet('收入金额');
  const COL_COUNT = 4;

  ws.columns = [
    { key: 'item', width: 22 },
    { key: 'qty', width: 16 },
    { key: 'amount', width: 22 },
    { key: 'tongzhou', width: 22 },
  ];

  ws.mergeCells('A1:D1');
  const titleCell = ws.getCell('A1');
  titleCell.value = '____年__月工作量、收入金额报表';
  titleCell.font = FONT_TITLE;
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getRow(1).height = 30;

  const deptRow = ws.getRow(2);
  deptRow.getCell(1).value = '科别：病理科';
  deptRow.font = FONT_HEADER;

  const headerRow = ws.getRow(3);
  headerRow.values = ['项目', '数量', '金额', '备注（通州）'];
  headerRow.font = FONT_HEADER;
  headerRow.eachCell(c => { c.fill = HEADER_FILL; c.border = BORDER_THIN; c.alignment = { horizontal: 'center' }; });

  const items = [
    { name: '病房', note: '填总金额和通州金额' },
    { name: '门诊', note: '填总金额和通州金额' },
    { name: '病理会诊（门诊）', note: '填 单价*数量=金额' },
    { name: '病理会诊（住院）', note: '填 单价*数量=金额' },
    { name: '门诊具体如下：', isSection: true },
  ];

  const priceTiers = [20, 30, 50, 80, 100, 120, 130, 140, 150, 920];
  priceTiers.forEach(p => items.push({ name: `  ${p}元档` }));
  items.push({ name: '合计', bold: true });

  items.forEach((item, idx) => {
    const row = ws.getRow(idx + 4);
    row.getCell(1).value = item.name;
    if (item.bold) row.getCell(1).font = FONT_HEADER;
    if (item.isSection) row.getCell(1).font = FONT_HEADER;
    applyBorders(row, COL_COUNT);
  });

  const noteStart = items.length + 5;
  const notes = [
    '说明：',
    '1. 金额栏格式：总金额（如 总6441870）',
    '2. 通州栏格式：通州金额（如 通2134570）',
    '3. 门诊具体部分每档填写：数量 和 金额',
    '4. 如有多个单价档位可在下方继续增加行',
  ];
  notes.forEach((note, i) => {
    ws.getRow(noteStart + i).getCell(1).value = note;
    ws.getRow(noteStart + i).getCell(1).font = { size: 9, color: { argb: 'FF666666' } };
  });
}

function buildConsultationSheet(wb) {
  const ws = wb.addWorksheet('会诊统计');
  const COL_COUNT = 4;

  ws.columns = [
    { key: 'item', width: 20 },
    { key: 'xizhimen', width: 24 },
    { key: 'tongzhou', width: 24 },
    { key: 'subtotal', width: 18 },
  ];

  ws.mergeCells('A1:D1');
  const titleCell = ws.getCell('A1');
  titleCell.value = '____年__月病理科会诊统计';
  titleCell.font = FONT_TITLE;
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getRow(1).height = 30;

  const headerRow = ws.getRow(2);
  headerRow.values = ['', '西直门', '通州', '合计'];
  headerRow.font = FONT_HEADER;
  headerRow.eachCell(c => { c.fill = HEADER_FILL; c.border = BORDER_THIN; c.alignment = { horizontal: 'center' }; });

  const items = [
    { name: '门诊会诊', note: '200元档：填 单价*数量=金额' },
    { name: '', note: '20元档：填 单价*数量=金额', hasSubtotal: true },
    { name: '住院会诊', note: '200元档：填 单价*数量=金额' },
    { name: '', note: '20元档：填 单价*数量=金额', hasSubtotal: true },
    { name: '共计', bold: true },
  ];

  items.forEach((item, idx) => {
    const row = ws.getRow(idx + 3);
    row.getCell(1).value = item.name;
    if (item.bold) row.font = FONT_HEADER;
    applyBorders(row, COL_COUNT);
  });

  const noteRow = ws.getRow(items.length + 4);
  noteRow.getCell(1).value = '填写格式：单价*数量=金额，如 200*510=102000';
  noteRow.getCell(1).font = { size: 9, color: { argb: 'FF666666' } };
}

function buildInstructionsSheet(wb) {
  const ws = wb.addWorksheet('填写说明');
  ws.getColumn(1).width = 80;

  const lines = [
    '【病理科月度工作报表模板 - 填写说明】',
    '',
    '一、工作量统计（Sheet 1）',
    '   - 填写各类检查的例数和块数',
    '   - 外检总量和通州外检需填写全部4列（例数、块数、冰冻例数、冰冻块数）',
    '   - 其他项目只需填写例数即可',
    '   - 第一行标题中填写年月，如"2026年4月病理科工作量"',
    '',
    '二、收入金额（Sheet 2）',
    '   - 病房和门诊行填写总金额和通州金额',
    '   - 病理会诊行填写各档位的 数量 和 金额',
    '   - 门诊具体部分按价位档填写数量和金额',
    '   - 合计行填写总合计金额',
    '',
    '三、会诊统计（Sheet 3）',
    '   - 按院区（西直门/通州）分列填写',
    '   - 每种会诊类型分200元档和20元档两行',
    '   - 填写格式：单价*数量=金额（如 200*510=102000）',
    '   - 系统会解析金额数字',
    '',
    '四、导入平台',
    '   - 保存文件后，在分析平台 → 月度工作报表 页面上传此文件',
    '   - 系统自动生成可视化图表',
    '   - 支持导出为 HTML 文件',
  ];

  lines.forEach((line, i) => {
    const row = ws.getRow(i + 1);
    row.getCell(1).value = line;
    if (i === 0) row.getCell(1).font = FONT_TITLE;
  });
}

if (require.main === module) {
  generateWorkloadTemplate().then(p => console.log('模板已生成:', p)).catch(console.error);
}

module.exports = { generateWorkloadTemplate };
