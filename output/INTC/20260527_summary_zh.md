# 回测分析报告 — INTC

**回测窗口**：2025-05-27 → 2026-05-26（共 251 个交易日）
**生成日期**：2026-05-27
**随机种子**：42 · **测试策略数量**：3

---

## 执行摘要

本次回测在英特尔（INTC）上测试了三种 MA/RSI 组合策略，覆盖完整一年的交易数据。三个策略均取得了可观的正收益，总回报率在 39.38% 至 43.82% 之间。MaRsiStrategy_Slow 以 1.513 的 Sharpe 比率位居风险调整收益第一，尽管其绝对收益排名第二。核心结论是：对于过去一年的 INTC 行情，慢速、保守的参数配置表现最优——Slow 策略以仅 8.14% 的最大回撤捕获了 41.64% 的总收益，风险收益比远优于其他两套策略。

## 策略对比

| 排名 | 策略 | 总收益率 | 年化收益 | 最大回撤 | Sharpe | 交易次数 | 胜率 |
|:----:|------|--------:|---------:|---------:|-------:|--------:|------:|
| 1 | MaRsiStrategy_Slow | 41.64% | 41.64% | 8.14% | 1.513 | 31 | 51.61% |
| 2 | MaRsiStrategy_Fast | 43.82% | 43.82% | 11.88% | 1.405 | 39 | 48.72% |
| 3 | MaRsiStrategy_Aggressive | 39.38% | 39.38% | 18.53% | 1.180 | 55 | 49.09% |

按 Sharpe（风险调整收益）降序排列。年化收益按实际数据窗口（251 个交易日，约一个自然年）折算。

## 风险标记

本次回测未触发任何关键风险标记。三个策略的最大回撤均控制在 20% 以下，且 Sharpe 比率均超过 1.0，整体表现优异。

## 部署建议

**推荐部署：MaRsiStrategy_Slow**

MaRsiStrategy_Slow 以 1.513 的 Sharpe 比率领先，明显优于 MaRsiStrategy_Fast（1.405）和 MaRsiStrategy_Aggressive（1.180）。其 8.14% 的最大回撤控制极为出色——不到 Aggressive 变体 18.53% 回撤的一半——适合风险偏好中等的账户。虽然 MaRsiStrategy_Fast 在绝对收益上以 2.18 个百分点微幅领先，但这一边际收益是以 46% 更大的回撤为代价换取的。Slow 策略全年仅 31 笔交易，交易成本和执行负担也更为可控。亚军 MaRsiStrategy_Fast 对于愿意承受更高回撤以换取略高收益的投资者而言，也是一个可行的备选方案。

## 产物清单

- **资金曲线对比图**：`output/INTC/20260527_comparison_chart.png`
- **逐笔交易统计（CSV）**：`output/INTC/20260527_trades.csv`
- **各策略 QuantStats 详细报告**：
  - `output/INTC/20260527_MaRsiStrategy_Slow_report.html`
  - `output/INTC/20260527_MaRsiStrategy_Fast_report.html`
  - `output/INTC/20260527_MaRsiStrategy_Aggressive_report.html`
- **完整运行日志**：`output/INTC/20260527_run.log`

QuantStats HTML 报告含完整的风险/收益指标（Sortino、Calmar、月度热力图、回撤明细），用浏览器打开即可深入分析。

---
*由 `backtrader-multi` skill 于 2026-05-27 生成。*
