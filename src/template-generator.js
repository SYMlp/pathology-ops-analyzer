const ExcelJS = require('exceljs');
const path = require('path');

const HEADER_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
const BORDER_THIN = {
  top: { style: 'thin' }, bottom: { style: 'thin' },
  left: { style: 'thin' }, right: { style: 'thin' },
};
const FONT_HEADER = { bold: true, size: 11 };
const FONT_TITLE = { bold: true, size: 14 };

async function generateTemplate(outputPath) {
  const wb = new ExcelJS.Workbook();
  wb.creator = '病理科运营分析系统';

  buildOperationsSheet(wb);
  buildRankingsSheet(wb);
  buildProcurementSheet(wb);
  buildDispatchSheet(wb);
  buildInstructionsSheet(wb);

  const filePath = outputPath || path.join(__dirname, '..', 'templates', '月度运营数据模板.xlsx');
  await wb.xlsx.writeFile(filePath);
  console.log(`模板已生成: ${filePath}`);
  return filePath;
}

function buildOperationsSheet(wb) {
  const ws = wb.addWorksheet('运营指标');

  const monthCols = [];
  for (let m = 1; m <= 12; m++) monthCols.push(`${m}月情况`);

  ws.columns = [
    { header: '', key: 'name', width: 28 },
    { header: '2025年度均值', key: 'prev_avg', width: 14 },
    { header: '本年度均值', key: 'cur_avg', width: 12 },
    ...monthCols.map((h, i) => ({ header: h, key: `m${i + 1}`, width: 11 })),
  ];

  ws.mergeCells('A1:O1');
  const titleCell = ws.getCell('A1');
  titleCell.value = '2026年__月病理科运营指标';
  titleCell.font = FONT_TITLE;
  titleCell.alignment = { horizontal: 'center' };

  const headerRow = ws.getRow(2);
  headerRow.values = ['运营指标\\提供数据周期', '2025年度均值', '本年度均值', ...monthCols];
  headerRow.font = FONT_HEADER;
  headerRow.eachCell(c => { c.border = BORDER_THIN; });

  const sections = [
    { label: '一、科室基本信息', isSection: true },
    { label: '科室人数' },
    { label: '二、科室经济指标', isSection: true },
    { label: '(一) 门诊收入', isSubsection: true },
    { label: '门诊计奖收入总计' },
    { label: '门诊医事服务费' },
    { label: '门诊运营收入总计', bold: true },
    { label: '门诊流水收入总计' },
    { label: '门诊材料费' },
    { label: '门诊药费' },
    { label: '(二) 科室支出', isSubsection: true },
    { label: '1. 售类消耗', isSubsub: true },
    { label: '非收费耗材', indent: true },
    { label: '办公耗材', indent: true },
    { label: '试剂领用', indent: true },
    { label: '医疗垃圾', indent: true },
    { label: '氧气', indent: true },
    { label: '设备类耗材', indent: true },
    { label: '2. 第三方费用' },
    { label: '3. 人员工资' },
    { label: '4. 兼职教授成本' },
    { label: '5. 设备折旧' },
    { label: '6. 设备维修费' },
    { label: '7. 基本奖金' },
    { label: '8. 夜班费' },
    { label: '(三) 科室运行情况汇总', isSubsection: true },
    { label: '科室计奖收入合计', bold: true },
    { label: '科室支出合计', bold: true },
    { label: '三、科室工作量指标', isSection: true },
    { label: '(一) 门诊工作量情况', isSubsection: true },
    { label: '专家门诊量' },
    { label: '特需门诊量' },
    { label: '普通门诊量' },
  ];

  sections.forEach((item, idx) => {
    const row = ws.getRow(idx + 3);
    row.getCell(1).value = item.indent ? `    ${item.label}` : item.label;

    if (item.isSection) {
      row.getCell(1).fill = HEADER_FILL;
      row.getCell(1).font = FONT_HEADER;
    }
    if (item.bold) {
      row.getCell(1).font = FONT_HEADER;
    }

    row.eachCell(c => { c.border = BORDER_THIN; });
  });

  const noteRow = ws.getRow(sections.length + 4);
  noteRow.getCell(1).value = '科室需注意事项：（人员经费+科室支出）/科室计奖收入';
  noteRow.getCell(1).font = { color: { argb: 'FFFF0000' } };

  const formulas = [
    '运营收入合计=门诊计奖收入+医事服务费(不含药耗)+病房计奖收入+医事服务费',
    '运营支出合计=人员支出+设备支出+非收费医材支出+其他消耗支出+第三方服务支出',
    '其他消耗支出=试剂+设备类耗材+办公耗材+医疗废弃物+氧气',
    '设备支出合计=设备折旧+设备维修',
  ];
  formulas.forEach((f, i) => {
    const r = ws.getRow(sections.length + 6 + i);
    r.getCell(1).value = f;
    r.getCell(1).font = { size: 9, color: { argb: 'FF666666' } };
  });
}

function buildRankingsSheet(wb) {
  const ws = wb.addWorksheet('全院排名');
  const headers = [
    '科室字典', '5分满意度', '满意度排名', '支收比',
    '支收比增量', '预约效率增量', '增量和(上四分位,下四分位)', '增量排序',
  ];

  ws.columns = headers.map(h => ({ header: h, width: 18 }));
  const headerRow = ws.getRow(1);
  headerRow.font = FONT_HEADER;
  headerRow.eachCell(c => { c.border = BORDER_THIN; });

  const dataRow = ws.getRow(2);
  dataRow.getCell(1).value = '病理科';
  dataRow.eachCell(c => { c.border = BORDER_THIN; });
}

function buildProcurementSheet(wb) {
  const ws = wb.addWorksheet('采购入库');
  const headers = [
    '订单号', '耗材名称', '规格', '供应商', '数量', '单价', '金额', '入库日期', '科室对应',
  ];

  ws.columns = [
    { header: headers[0], width: 16 },
    { header: headers[1], width: 28 },
    { header: headers[2], width: 20 },
    { header: headers[3], width: 20 },
    { header: headers[4], width: 10 },
    { header: headers[5], width: 12 },
    { header: headers[6], width: 14 },
    { header: headers[7], width: 14 },
    { header: headers[8], width: 14 },
  ];

  const headerRow = ws.getRow(1);
  headerRow.font = FONT_HEADER;
  headerRow.eachCell(c => { c.border = BORDER_THIN; });
}

function buildDispatchSheet(wb) {
  const ws = wb.addWorksheet('出库领用');
  const headers = [
    '序号', '物品名称', '规格', '出库科室', '收费科室',
    '数量', '金额', '使用日期', '生产厂家', '科室对应',
  ];

  ws.columns = [
    { header: headers[0], width: 8 },
    { header: headers[1], width: 28 },
    { header: headers[2], width: 20 },
    { header: headers[3], width: 14 },
    { header: headers[4], width: 14 },
    { header: headers[5], width: 10 },
    { header: headers[6], width: 14 },
    { header: headers[7], width: 14 },
    { header: headers[8], width: 20 },
    { header: headers[9], width: 14 },
  ];

  const headerRow = ws.getRow(1);
  headerRow.font = FONT_HEADER;
  headerRow.eachCell(c => { c.border = BORDER_THIN; });
}

function buildInstructionsSheet(wb) {
  const ws = wb.addWorksheet('填写说明');
  ws.getColumn(1).width = 80;

  const lines = [
    '【病理科月度运营数据模板 - 填写说明】',
    '',
    '一、运营指标（Sheet 1）',
    '   - 每月填写对应月份列的数据',
    '   - "2025年度均值" 和 "本年度均值" 根据历史数据填写',
    '   - 黄色底色行为分类标题行，无需填写数据',
    '   - 缩进行（如非收费耗材）为子类目，对应大类',
    '   - 加粗行为汇总行',
    '',
    '二、全院排名（Sheet 2）',
    '   - 每月更新排名数据',
    '   - 满意度为5分制，保留3位小数',
    '   - 支收比 = 支出/收入，保留3位小数',
    '',
    '三、采购入库（Sheet 3）',
    '   - 每行一笔采购入库记录',
    '   - 金额单位为元',
    '   - 日期格式：YYYY-MM-DD 或 YYYY/MM/DD',
    '',
    '四、出库领用（Sheet 4）',
    '   - 每行一笔出库记录',
    '   - 金额单位为元',
    '   - 日期格式同上',
    '',
    '五、导入平台',
    '   - 保存文件后，在分析平台上传此文件',
    '   - 系统自动识别各 Sheet 的数据并生成分析报告',
    '   - 导出的 HTML 报告可直接在浏览器打开查看',
  ];

  lines.forEach((line, i) => {
    const row = ws.getRow(i + 1);
    row.getCell(1).value = line;
    if (i === 0) row.getCell(1).font = FONT_TITLE;
  });
}

if (require.main === module) {
  generateTemplate().catch(console.error);
}

module.exports = { generateTemplate };
