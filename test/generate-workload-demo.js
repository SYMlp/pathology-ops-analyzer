const ExcelJS = require('exceljs');
const path = require('path');

const BORDER = {
  top: { style: 'thin' }, bottom: { style: 'thin' },
  left: { style: 'thin' }, right: { style: 'thin' },
};
const HEAD_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
const F_HEAD = { bold: true, size: 11 };
const F_TITLE = { bold: true, size: 14 };

async function buildDemo(outPath) {
  const wb = new ExcelJS.Workbook();

  const ws1 = wb.addWorksheet('工作量统计');
  ws1.columns = [
    { width: 18 }, { width: 14 }, { width: 14 }, { width: 14 }, { width: 14 },
  ];
  ws1.mergeCells('A1:E1');
  ws1.getCell('A1').value = '2026年4月病理科工作量';
  ws1.getCell('A1').font = F_TITLE;
  ws1.getCell('A1').alignment = { horizontal: 'center' };
  ws1.getRow(1).height = 28;
  ws1.getRow(2).values = ['名称', '例数', '块数', '冰冻例数', '冰冻块数'];
  ws1.getRow(2).font = F_HEAD;
  ws1.getRow(2).eachCell(c => { c.fill = HEAD_FILL; c.border = BORDER; c.alignment = { horizontal: 'center' }; });

  const rows = [
    ['外检总量', 13005, 83234, 1021, 2578],
    ['通州外检', 2192, 23558, 233, 539],
    ['妇科细胞审核', 2412, null, null, null],
    ['肾内切片审核', 172, null, null, null],
    ['分子病理', 105, null, null, null],
    ['基因检测', 1024, null, null, null],
  ];
  rows.forEach((r, i) => {
    const row = ws1.getRow(i + 3);
    r.forEach((v, c) => { row.getCell(c + 1).value = v; });
    row.eachCell({ includeEmpty: true }, c => { c.border = BORDER; });
  });

  const ws2 = wb.addWorksheet('收入金额');
  ws2.columns = [{ width: 22 }, { width: 16 }, { width: 22 }, { width: 22 }];
  ws2.mergeCells('A1:D1');
  ws2.getCell('A1').value = '2026年4月工作量、收入金额报表';
  ws2.getCell('A1').font = F_TITLE;
  ws2.getCell('A1').alignment = { horizontal: 'center' };
  ws2.getRow(1).height = 28;
  ws2.getRow(2).getCell(1).value = '科别：病理科';
  ws2.getRow(2).getCell(1).font = F_HEAD;
  ws2.getRow(3).values = ['项目', '数量', '金额', '备注（通州）'];
  ws2.getRow(3).font = F_HEAD;
  ws2.getRow(3).eachCell(c => { c.fill = HEAD_FILL; c.border = BORDER; c.alignment = { horizontal: 'center' }; });

  const r2 = [
    ['病房', null, '总6441870', '通2134570'],
    ['门诊', null, '总3453930', '通202170'],
    ['病理会诊（门诊）', null, '总200*535=107000', '通200*25=5000'],
    ['病理会诊（住院）', null, '总200*118=23600', '通200*41=8200'],
    ['门诊具体如下：', null, null, null],
    ['20元档', 5631, 112620, '通州'],
    ['30元档', 1770, 53100, null],
    ['50元档', 3860, 193000, null],
    ['80元档', 93, 7440, null],
    ['100元档', 2, 200, null],
    ['120元档', 135, 16200, null],
    ['130元档', 105, 13650, null],
    ['140元档', 165, 23100, null],
    ['150元档', 16318, 2447700, null],
    ['920元档', 2, 1840, null],
    ['合计', null, '总9895800', '通2336740'],
  ];
  r2.forEach((r, i) => {
    const row = ws2.getRow(i + 4);
    r.forEach((v, c) => { row.getCell(c + 1).value = v; });
    row.eachCell({ includeEmpty: true }, c => { c.border = BORDER; });
  });

  const ws3 = wb.addWorksheet('会诊统计');
  ws3.columns = [{ width: 20 }, { width: 24 }, { width: 24 }, { width: 18 }];
  ws3.mergeCells('A1:D1');
  ws3.getCell('A1').value = '2026年4月病理科会诊统计';
  ws3.getCell('A1').font = F_TITLE;
  ws3.getCell('A1').alignment = { horizontal: 'center' };
  ws3.getRow(1).height = 28;
  ws3.getRow(2).values = ['', '西直门', '通州', '合计'];
  ws3.getRow(2).font = F_HEAD;
  ws3.getRow(2).eachCell(c => { c.fill = HEAD_FILL; c.border = BORDER; c.alignment = { horizontal: 'center' }; });

  const r3 = [
    ['门诊会诊', '200*510=102000', '200*25=5000', null],
    ['', '20*4475=89500', '20*476=9520', 206020],
    ['住院会诊', '200*77=15400', '200*41=8200', null],
    ['', '20*352=7040', '20*326=6520', 37160],
    ['共计', 213940, 29240, 243180],
  ];
  r3.forEach((r, i) => {
    const row = ws3.getRow(i + 3);
    r.forEach((v, c) => { row.getCell(c + 1).value = v; });
    row.eachCell({ includeEmpty: true }, c => { c.border = BORDER; });
  });

  await wb.xlsx.writeFile(outPath);
  console.log('demo Excel written:', outPath);
}

if (require.main === module) {
  const out = path.join(__dirname, '..', 'uploads', '_demo-workload-2026-04.xlsx');
  buildDemo(out).catch(e => { console.error(e); process.exit(1); });
}

module.exports = { buildDemo };
