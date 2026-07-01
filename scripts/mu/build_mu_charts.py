"""
Micron Technology (MU) Q3 FY2026 Earnings Update -- Chart Generator.
Quarter ended: May 28, 2026 | Reported: June 24, 2026

Charts are intentionally generated with Pillow instead of matplotlib so the
script works with the bundled Codex Python runtime.
"""

from mu_q3_fy2026_common import build_all_charts


if __name__ == "__main__":
    paths = build_all_charts()
    for path in paths:
        print(f"Saved: {path}")
    print(f"\nAll {len(paths)} MU Q3 FY2026 charts generated successfully.")
