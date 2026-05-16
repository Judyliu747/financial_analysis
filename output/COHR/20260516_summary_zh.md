# 回测分析报告 — COHR

**回测窗口**：2025-05-16 → 2026-05-15（共 251 个交易日）
**生成日期**：2026-05-16
**随机种子**：42 · **测试策略数量**：3

---

## 执行摘要

本次回测在 Coherent Corp（COHR）上测试了三种 MA/RSI 组合策略，覆盖完整一年的交易数据。MaRsiStrategy_Aggressive 以 30.10% 的总收益率和 0.993 的 Sharpe 比率位居第一，但触发了高回撤标记，最大回撤达 22.24%。核心结论是：更快、更激进的参数配置能更好地捕捉 COHR 的动量驱动行情，但伴随的高回撤要求部署者具备充足的风险承受能力。

## 策略对比

| 排名 | 策略 | 总收益率 | 年化收益 | 最大回撤 | Sharpe | 交易次数 | 胜率 |
|:----:|------|--------:|---------:|---------:|-------:|--------:|------:|
| 1 | MaRsiStrategy_Aggressive | 30.10% | 30.10% | 22.24% | 0.993 | 61 | 55.74% |
| 2 | MaRsiStrategy_Fast | 24.85% | 24.85% | 19.46% | 0.853 | 43 | 62.79% |
| 3 | MaRsiStrategy_Slow | -0.99% | -0.99% | 17.88% | 0.082 | 22 | 59.09% |

按 Sharpe（风险调整收益）降序排列。年化收益按实际数据窗口（251 个交易日，约一个自然年）折算。

## 风险标记

- **MaRsiStrategy_Aggressive** — `high_drawdown`：最大回撤达 22.24%，超过 20% 警戒线。部署该策略需要能够承受超过五分之一的账户净值从峰值到谷底的下行，且不会触发强制平仓或情绪化退出。

## 部署建议

**推荐部署：MaRsiStrategy_Fast**

尽管 MaRsiStrategy_Aggressive 取得了最高的绝对收益和风险调整收益，但其 22.24% 的最大回撤在实盘部署中存在隐患——尤其是对于缺乏深度回撤缓冲的账户。MaRsiStrategy_Fast 实现了 24.85% 的收益率，Sharpe 为 0.853，同时将最大回撤控制在 19.46%，刚好低于 20% 的警戒线。此外，该策略在 43 笔交易中取得了三者最高的 62.79% 胜率，表明其信号质量更为稳定。与 Aggressive 变体的差距确实存在（收益率差 5.25 个百分点，Sharpe 差 0.140），但更干净的风险特征使 Fast 成为初始部署时更审慎的选择。建议将两者同时在模拟盘中运行，验证实盘条件下的回撤差异后再考虑切换至 Aggressive。

## 产物清单

- **资金曲线对比图**：`output/COHR/20260516_comparison_chart.png`
- **逐笔交易统计（CSV）**：`output/COHR/20260516_trades.csv`
- **各策略 QuantStats 详细报告**：
  - `output/COHR/20260516_MaRsiStrategy_Aggressive_report.html`
  - `output/COHR/20260516_MaRsiStrategy_Fast_report.html`
  - `output/COHR/20260516_MaRsiStrategy_Slow_report.html`
- **完整运行日志**：`output/COHR/20260516_run.log`

QuantStats HTML 报告含完整的风险/收益指标（Sortino、Calmar、月度热力图、回撤明细），用浏览器打开即可深入分析。

---
*由 `backtrader-multi` skill 于 2026-05-16 生成。*
