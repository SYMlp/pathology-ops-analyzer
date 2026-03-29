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

```bash
npm install
npm start          # 启动平台 http://localhost:3200
npm run generate-template  # 单独生成 Excel 模板
```

## 部署到 Vercel

```bash
npm i -g vercel
cd pathology-ops-analyzer
vercel
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
