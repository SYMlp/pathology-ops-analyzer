const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const os = require('os');
const dayjs = require('dayjs');

const { parseExcel } = require('./src/parser');
const { analyze, postAnalyze } = require('./src/analyzer');
const { generateReport, generateReportBody } = require('./src/reporter');
const { generateTemplate } = require('./src/template-generator');

const app = express();
const PORT = process.env.PORT || 3200;

app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));
app.use(express.static(path.join(__dirname, 'public')));

const uploadDir = process.env.VERCEL ? os.tmpdir() : path.join(__dirname, 'uploads');

const upload = multer({
  dest: uploadDir,
  limits: { fileSize: 20 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (['.xlsx', '.xls'].includes(ext)) cb(null, true);
    else cb(new Error('仅支持 .xlsx 或 .xls 格式'));
  },
});

let latestAnalysis = null;
app.get('/favicon.ico', (req, res) => res.status(204).end());
app.get('/', (req, res) => {
  let reportBody = null;
  if (latestAnalysis) reportBody = generateReportBody(latestAnalysis);
  res.render('index', { analysis: latestAnalysis, reportBody, error: null });
});

app.post('/upload', upload.single('datafile'), async (req, res) => {
  try {
    if (!req.file) throw new Error('请选择文件上传');

    const filePath = req.file.path;
    const parsed = await parseExcel(filePath);
    const result = analyze(parsed);
    postAnalyze(result);
    latestAnalysis = result;

    try { fs.unlinkSync(filePath); } catch (_) {}

    res.redirect('/');
  } catch (err) {
    if (req.file?.path) {
      try { fs.unlinkSync(req.file.path); } catch (_) {}
    }
    res.render('index', { analysis: latestAnalysis, reportBody: latestAnalysis ? generateReportBody(latestAnalysis) : null, error: err.message });
  }
});

app.get('/export', (req, res) => {
  if (!latestAnalysis) return res.redirect('/');

  const html = generateReport(latestAnalysis);
  const { year, month, department } = latestAnalysis.meta;
  const filename = `${department}运营分析_${year}年${month}月_${dayjs().format('MMDD')}.html`;

  res.setHeader('Content-Type', 'text/html; charset=utf-8');
  res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`);
  res.send(html);
});

app.get('/preview', (req, res) => {
  if (!latestAnalysis) return res.redirect('/');
  const html = generateReport(latestAnalysis);
  res.send(html);
});

app.get('/pdf', async (req, res) => {
  if (!latestAnalysis) return res.redirect('/');
  try {
    const puppeteer = require('puppeteer');
    const body = generateReportBody(latestAnalysis);
    const html = `<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Noto+Sans+SC:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
body{background:#fff;margin:0}
details{display:block!important}
details>summary{display:none!important}
details .fold-body,details .insight-fold-body{display:block!important;padding:8px 0}
.btn-action,.header-actions{display:none!important}
</style>
</head><body>${body}</body></html>`;

    const browser = await puppeteer.launch({ headless: true, args: ['--no-sandbox', '--disable-setuid-sandbox'] });
    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: 'networkidle0', timeout: 30000 });
    // Wait for ECharts to render
    await new Promise(r => setTimeout(r, 2000));

    const pdfBuffer = await page.pdf({
      format: 'A4',
      landscape: true,
      margin: { top: '8mm', right: '8mm', bottom: '8mm', left: '8mm' },
      printBackground: true,
    });
    await browser.close();

    const { year, month, department } = latestAnalysis.meta;
    const filename = `${department}运营分析_${year}年${month}月.pdf`;
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`);
    res.send(pdfBuffer);
  } catch (err) {
    console.error('PDF generation error:', err);
    res.status(500).send('PDF 生成失败: ' + err.message);
  }
});

app.get('/template', async (req, res) => {
  try {
    const templateDir = process.env.VERCEL ? os.tmpdir() : path.join(__dirname, 'templates');
    const filePath = path.join(templateDir, '月度运营数据模板.xlsx');
    
    // Always regenerate to stay up to date
    await generateTemplate(filePath);
    res.download(filePath, '病理科月度运营数据模板.xlsx');
  } catch (err) {
    res.status(500).send('模板生成失败: ' + err.message);
  }
});

app.get('/api/analysis', (req, res) => {
  if (!latestAnalysis) return res.status(404).json({ error: '暂无分析数据' });
  res.json(latestAnalysis);
});

if (!process.env.VERCEL) {
  app.listen(PORT, () => {
    console.log(`\n  科室运营分析平台已启动`);
    console.log(`  访问地址: http://localhost:${PORT}`);
    console.log(`  模板下载: http://localhost:${PORT}/template\n`);
  });
}

module.exports = app;
