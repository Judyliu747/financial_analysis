# 回测分析报告 — RKLB

**回测窗口**：2025-06-24 → 2026-06-23（共 251 个交易日）
**生成日期**：2026-06-24
**随机种子**：42 · **测试策略数量**：3

---

## 执行摘要

本次回测对 Rocket Lab USA (RKLB) 在过去一年内的交易数据运行了三种 MA+RSI 组合策略——MaRsiStrategy_Slow、MaRsiStrategy_Fast 和 MaRsiStrategy_Aggressive。按风险调整收益（Sharpe）排名，MaRsiStrategy_Aggressive 以 1.681 的 Sharpe 比率位列第一，总收益率达 76.07%，同时将最大回撤控制在 18.69%。整体来看，三种策略在 RKLB 的强势上涨行情中均实现了正收益，但激进策略在收益和风险控制之间取得了最优的平衡。

## 策略对比

| 排名 | 策略 | 总收益率 | 年化收益 | 最大回撤 | Sharpe | 交易次数 | 胜率 |
|:----:|------|--------:|---------:|---------:|-------:|--------:|------:|
| 1 | MaRsiStrategy_Aggressive | 76.07% | 76.49% | 18.69% | 1.681 | 63 | 58.73% |
| 2 | MaRsiStrategy_Slow | 46.14% | 46.41% | 8.66% | 1.526 | 41 | 60.98% |
| 3 | MaRsiStrategy_Fast | 29.37% | 29.52% | 19.74% | 0.884 | 50 | 58.00% |

按 Sharpe（风险调整收益）降序排列。年化收益按实际数据窗口（251 个交易日）折算。

## 风险标记

本次回测未触发任何关键风险标记。三种策略的最大回撤均未超过 20% 阈值，Sharpe 比率全部为正，且交易样本量充足（均超过 40 次）。

值得注意的是，MaRsiStrategy_Fast 的最大回撤为 19.74%，已接近 20% 的高回撤警戒线，在实盘部署时应额外关注资金管理。

## 部署建议

**推荐部署：MaRsiStrategy_Aggressive**

MaRsiStrategy_Aggressive 凭借 1.681 的 Sharpe 比率在三种策略中排名第一，同时实现了 76.07% 的总收益率——远超第二名 MaRsiStrategy_Slow 的 46.14%。虽然其最大回撤（18.69%）高于 MaRsiStrategy_Slow 的 8.66%，但仍控制在 20% 以内，处于可接受范围。该策略在 63 笔交易中保持了 58.73% 的胜率，单笔盈利均值（约 48.6 万元）高于单笔亏损均值（约 40.9 万元），呈现正的盈亏比。

需要指出的是，MaRsiStrategy_Slow 虽然 Sharpe 略低（1.526），但其仅 8.66% 的最大回撤和 60.98% 的胜率使其成为风险厌恶型投资者的理想备选。如果资金对波动的承受能力有限，MaRsiStrategy_Slow 是更稳健的选择。

## 产物清单

- **资金曲线对比图**：`output/RKLB/20260624_comparison_chart.png`
- **逐笔交易统计（CSV）**：`output/RKLB/20260624_trades.csv`
- **各策略 QuantStats 详细报告**：
  - `output/RKLB/20260624_MaRsiStrategy_Aggressive_report.html`
  - `output/RKLB/20260624_MaRsiStrategy_Fast_report.html`
  - `output/RKLB/20260624_MaRsiStrategy_Slow_report.html`
- **完整运行日志**：`output/RKLB/20260624_run.log`

QuantStats HTML 报告含完整的风险/收益指标（Sortino、Calmar、月度热力图、回撤明细），用浏览器打开即可深入分析。

---
*由 `backtrader-multi` skill 于 2026-06-24 生成。*
