# 科室运营分析平台

病理科月度运营数据分析工具。上传 Excel 数据，自动生成专业级可视化分析报告。

## 功能

- **经济运行分析** — 收入支出趋势、支出结构、同比环比
- **运营效率指标** — 支收比、人力成本占比、百元收入卫材消耗
- **工作量分析** — 门诊量趋势、人均工作量
- **全院排名** — 满意度、支收比增量、综合排名
- **耗材采购** — 采购 TOP10、供应商分布、出库明细
- **智能洞察** — 风险预警、运营亮点、改进建议

## 使用

门户首页为 `/`，可从两个入口进入：**科室运营分析**（`/ops`，上传运营大表）与 **工作报表柱状图**（`/workload`，专用工作报表模板）。

```bash
npm install
npm start          # 门户 http://localhost:3200 · 科室运营 http://localhost:3200/ops
npm run generate-template  # 单独生成 Excel 模板
```

## 部署到 Vercel

**线上地址：** https://syw.lsrabbit.space （门户 `/`，科室运营 `/ops`，工作报表 `/workload`）

```bash
npm i -g vercel
cd pathology-ops-analyzer
vercel --prod
```

## Excel 模板

4 个 Sheet：

| Sheet | 内容 |
|:---|:---|
| 运营指标 | 收入、支出、工作量月度数据 |
| 全院排名 | 满意度、支收比等排名 |
| 采购入库 | 逐笔采购记录 |
| 出库领用 | 逐笔出库记录 |

## 技术栈

Node.js + Express + EJS + ExcelJS + Chart.js
