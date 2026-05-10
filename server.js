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
const { generateWorkloadTemplate } = require('./src/workload-template');
const { parseWorkloadExcel } = require('./src/workload-parser');
const { generateWorkloadReport, generateWorkloadReportBody } = require('./src/workload-reporter');
const { parseCompareExcel } = require('./src/compare-parser');
const { generateCompareReport, generateCompareReportBody } = require('./src/compare-reporter');

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
let latestWorkload = null;
let latestCompare = null;

const STATE_DIR = process.env.VERCEL ? os.tmpdir() : path.join(__dirname, 'uploads');
const ANALYSIS_STATE_FILE = path.join(STATE_DIR, '_analysis-state.json');
const WORKLOAD_STATE_FILE = path.join(STATE_DIR, '_workload-state.json');
const COMPARE_STATE_FILE = path.join(STATE_DIR, '_compare-state.json');

function persistAnalysis(data) {
  try {
    fs.writeFileSync(ANALYSIS_STATE_FILE, JSON.stringify(data));
  } catch (e) {
    console.warn('persist analysis state failed:', e.message);
  }
}

function loadAnalysis() {
  if (latestAnalysis) return latestAnalysis;
  try {
    if (fs.existsSync(ANALYSIS_STATE_FILE)) {
      const raw = fs.readFileSync(ANALYSIS_STATE_FILE, 'utf8');
      latestAnalysis = JSON.parse(raw);
      return latestAnalysis;
    }
  } catch (e) {
    console.warn('load analysis state failed:', e.message);
  }
  return null;
}

function persistWorkload(data) {
  try {
    fs.writeFileSync(WORKLOAD_STATE_FILE, JSON.stringify(data));
  } catch (e) {
    console.warn('persist workload state failed:', e.message);
  }
}

function loadWorkload() {
  if (latestWorkload) return latestWorkload;
  try {
    if (fs.existsSync(WORKLOAD_STATE_FILE)) {
      const raw = fs.readFileSync(WORKLOAD_STATE_FILE, 'utf8');
      latestWorkload = JSON.parse(raw);
      return latestWorkload;
    }
  } catch (e) {
    console.warn('load workload state failed:', e.message);
  }
  return null;
}

function persistCompare(data) {
  try {
    fs.writeFileSync(COMPARE_STATE_FILE, JSON.stringify(data));
  } catch (e) {
    console.warn('persist compare state failed:', e.message);
  }
}

function loadCompare() {
  if (latestCompare) return latestCompare;
  try {
    if (fs.existsSync(COMPARE_STATE_FILE)) {
      const raw = fs.readFileSync(COMPARE_STATE_FILE, 'utf8');
      latestCompare = JSON.parse(raw);
      return latestCompare;
    }
  } catch (e) {
    console.warn('load compare state failed:', e.message);
  }
  return null;
}
app.get('/favicon.ico', (req, res) => res.status(204).end());
app.get('/', (req, res) => {
  res.render('home');
});

app.get('/ops', (req, res) => {
  const analysis = loadAnalysis();
  const reportBody = analysis ? generateReportBody(analysis) : null;
  res.render('ops', { analysis, reportBody, error: null });
});

app.post('/upload', upload.single('datafile'), async (req, res) => {
  try {
    if (!req.file) throw new Error('请选择文件上传');

    const filePath = req.file.path;
    const parsed = await parseExcel(filePath);
    const result = analyze(parsed);
    postAnalyze(result);
    latestAnalysis = result;
    persistAnalysis(result);

    try { fs.unlinkSync(filePath); } catch (_) {}

    res.redirect('/ops');
  } catch (err) {
    if (req.file?.path) {
      try { fs.unlinkSync(req.file.path); } catch (_) {}
    }
    const analysis = loadAnalysis();
    res.render('ops', { analysis, reportBody: analysis ? generateReportBody(analysis) : null, error: err.message });
  }
});

app.get('/export', (req, res) => {
  const analysis = loadAnalysis();
  if (!analysis) return res.redirect('/ops');

  const html = generateReport(analysis);
  const { year, month, department } = analysis.meta;
  const filename = `${department}运营分析_${year}年${month}月_${dayjs().format('MMDD')}.html`;

  res.setHeader('Content-Type', 'text/html; charset=utf-8');
  res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`);
  res.send(html);
});

app.get('/preview', (req, res) => {
  const analysis = loadAnalysis();
  if (!analysis) return res.redirect('/ops');
  const html = generateReport(analysis);
  res.send(html);
});

app.get('/pdf', async (req, res) => {
  const analysis = loadAnalysis();
  if (!analysis) return res.redirect('/ops');
  try {
    let browser;
    if (process.env.VERCEL) {
      const chromium = require('@sparticuz/chromium');
      const puppeteer = require('puppeteer-core');
      browser = await puppeteer.launch({
        args: chromium.args,
        defaultViewport: chromium.defaultViewport,
        executablePath: await chromium.executablePath(),
        headless: chromium.headless,
      });
    } else {
      const puppeteer = require('puppeteer');
      browser = await puppeteer.launch({ headless: true, args: ['--no-sandbox', '--disable-setuid-sandbox'] });
    }

    const page = await browser.newPage();
    await page.setViewport({ width: 1200, height: 800 });

    const html = generateReport(analysis);
    await page.setContent(html, { waitUntil: 'networkidle0', timeout: 30000 });

    await page.evaluate(() => {
      document.querySelectorAll('details').forEach(d => d.setAttribute('open', ''));
    });

    await page.addStyleTag({ content: `
      .btn-action, .header-actions, .report-header .btn-action { display:none!important }
      .report { max-width:100%!important; padding:12px 20px 24px!important; margin:0!important }
    `});

    await new Promise(r => setTimeout(r, 2000));

    const pdfBuffer = await page.pdf({
      format: 'A4',
      landscape: true,
      margin: { top: '6mm', right: '6mm', bottom: '6mm', left: '6mm' },
      printBackground: true,
    });
    await browser.close();

    const { year, month, department } = analysis.meta;
    const filename = `${department}运营分析_${year}年${month}月.pdf`;
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Length', pdfBuffer.length);
    res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`);
    res.end(pdfBuffer);
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
  const analysis = loadAnalysis();
  if (!analysis) return res.status(404).json({ error: '暂无分析数据' });
  res.json(analysis);
});

app.get('/workload', (req, res) => {
  const data = loadWorkload();
  const workloadReportBody = data ? generateWorkloadReportBody(data, { downloadButtons: true }) : null;
  res.render('workload', { workloadAnalysis: data, workloadReportBody, error: null });
});

app.post('/workload/upload', upload.single('datafile'), async (req, res) => {
  try {
    if (!req.file) throw new Error('请选择文件上传');
    const filePath = req.file.path;
    const parsed = await parseWorkloadExcel(filePath);
    latestWorkload = parsed;
    persistWorkload(parsed);
    try {
      fs.unlinkSync(filePath);
    } catch (_) {}
    res.redirect('/workload');
  } catch (err) {
    if (req.file?.path) {
      try {
        fs.unlinkSync(req.file.path);
      } catch (_) {}
    }
    const data = loadWorkload();
    const workloadReportBody = data ? generateWorkloadReportBody(data, { downloadButtons: true }) : null;
    res.render('workload', { workloadAnalysis: data, workloadReportBody, error: err.message });
  }
});

app.get('/workload/template', async (req, res) => {
  try {
    const templateDir = process.env.VERCEL ? os.tmpdir() : path.join(__dirname, 'templates');
    const filePath = path.join(templateDir, '月度工作报表模板.xlsx');
    await generateWorkloadTemplate(filePath);
    res.download(filePath, '病理科月度工作报表模板.xlsx');
  } catch (err) {
    res.status(500).send('工作报表模板生成失败: ' + err.message);
  }
});

app.get('/workload/export', (req, res) => {
  const data = loadWorkload();
  if (!data) return res.redirect('/workload');
  const html = generateWorkloadReport(data);
  const period = (data.meta && data.meta.period) || '';
  const filename = `病理科工作报表_${period || dayjs().format('YYYY-MM')}_${dayjs().format('MMDD')}.html`;
  res.setHeader('Content-Type', 'text/html; charset=utf-8');
  res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`);
  res.send(html);
});

app.get('/workload/preview', (req, res) => {
  const data = loadWorkload();
  if (!data) return res.redirect('/workload');
  res.send(generateWorkloadReport(data));
});

app.get('/workload/pdf', async (req, res) => {
  const data = loadWorkload();
  if (!data) return res.redirect('/workload');
  try {
    let browser;
    if (process.env.VERCEL) {
      const chromium = require('@sparticuz/chromium');
      const puppeteer = require('puppeteer-core');
      browser = await puppeteer.launch({
        args: chromium.args,
        defaultViewport: chromium.defaultViewport,
        executablePath: await chromium.executablePath(),
        headless: chromium.headless,
      });
    } else {
      const puppeteer = require('puppeteer');
      browser = await puppeteer.launch({ headless: true, args: ['--no-sandbox', '--disable-setuid-sandbox'] });
    }
    const page = await browser.newPage();
    await page.setViewport({ width: 1280, height: 1800 });
    const html = generateWorkloadReport(data);
    await page.setContent(html, { waitUntil: 'networkidle0', timeout: 30000 });
    await new Promise(r => setTimeout(r, 1500));
    const pdfBuffer = await page.pdf({
      format: 'A4',
      printBackground: true,
      margin: { top: '8mm', right: '8mm', bottom: '8mm', left: '8mm' },
    });
    await browser.close();
    const period = (data.meta && data.meta.period) || '';
    const filename = `病理科工作报表_${period || 'export'}.pdf`;
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Length', pdfBuffer.length);
    res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`);
    res.end(pdfBuffer);
  } catch (err) {
    console.error('Workload PDF error:', err);
    res.status(500).send('PDF 生成失败: ' + err.message);
  }
});

app.get('/compare', (req, res) => {
  const data = loadCompare();
  const compareReportBody = data ? generateCompareReportBody(data) : null;
  res.render('compare', { compareData: data, compareReportBody, error: null });
});

app.post('/compare/upload', upload.array('datafiles', 24), async (req, res) => {
  try {
    const files = req.files || [];
    if (files.length < 2) {
      throw new Error('至少需要上传 2 个月份的 Excel 文件才能做对比');
    }

    const monthly = [];
    for (const f of files) {
      try {
        const parsed = await parseCompareExcel(f.path);
        if (!parsed.meta || !parsed.meta.sortKey) {
          throw new Error(`文件 ${f.originalname} 无法识别月份（请检查表格首行是否含"YYYY年M月"）`);
        }
        monthly.push(parsed);
      } finally {
        try { fs.unlinkSync(f.path); } catch (_) {}
      }
    }
    monthly.sort((a, b) => a.meta.sortKey - b.meta.sortKey);

    latestCompare = monthly;
    persistCompare(monthly);
    res.redirect('/compare');
  } catch (err) {
    if (req.files) {
      for (const f of req.files) {
        try { fs.unlinkSync(f.path); } catch (_) {}
      }
    }
    const data = loadCompare();
    res.render('compare', {
      compareData: data,
      compareReportBody: data ? generateCompareReportBody(data) : null,
      error: err.message,
    });
  }
});

app.get('/compare/preview', (req, res) => {
  const data = loadCompare();
  if (!data) return res.redirect('/compare');
  res.send(generateCompareReport(data));
});

app.get('/compare/export', (req, res) => {
  const data = loadCompare();
  if (!data) return res.redirect('/compare');
  const html = generateCompareReport(data);
  const first = data[0].meta.period;
  const last = data[data.length - 1].meta.period;
  const filename = `病理科多月对比_${first}-${last}.html`;
  res.setHeader('Content-Type', 'text/html; charset=utf-8');
  res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`);
  res.send(html);
});

if (!process.env.VERCEL) {
  app.listen(PORT, () => {
    console.log(`\n  病理数据平台已启动`);
    console.log(`  门户首页: http://localhost:${PORT}/`);
    console.log(`  科室运营: http://localhost:${PORT}/ops`);
    console.log(`  运营模板: http://localhost:${PORT}/template`);
    console.log(`  工作报表: http://localhost:${PORT}/workload`);
    console.log(`  多月对比: http://localhost:${PORT}/compare\n`);
  });
}

module.exports = app;
