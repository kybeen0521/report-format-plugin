"""예제 그림 2장(형식 확인용, 가상 자료). 실행: python3 make_example_figs.py [출력폴더]  (기본: 이 파일 옆)"""
import os, sys
import numpy as np, matplotlib
OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
matplotlib.use("Agg"); import matplotlib.pyplot as plt
rng = np.random.default_rng(0)
# 그림 1: 조건별 곡선(선 종류로 구분, 흑백)
fig, axs = plt.subplots(1, 2, figsize=(7.6, 3.1))
ls = {"A": ":", "B": "--", "C": "-"}
x = np.linspace(0, 10, 12)
for k, (ax, title) in enumerate(zip(axs, ("(a) case 1", "(b) case 2"))):
    for name, s in ls.items():
        y = 50 + 4 * (ord(name) - 64) * np.exp(-x / (4 + k)) + rng.normal(0, 0.6, x.size)
        ax.plot(x, y, ls=s, marker="o", ms=3, mfc="w", color="k", lw=1, label=f"setting {name}")
    ax.set_title(title, fontsize=10, fontweight="bold", loc="center"); ax.set_xlabel("x [unit]"); ax.set_ylabel("y [unit]"); ax.grid(ls=":", lw=.4)
axs[0].legend(fontsize=7, frameon=False)
fig.tight_layout(); fig.savefig(os.path.join(OUT, "fig1_curves.png"), dpi=160, bbox_inches="tight", pad_inches=0.06); plt.close(fig)
# 그림 2: 2+1 배치 등고선 3장, 아랫줄 빈 칸에 색막대·범례
from matplotlib.lines import Line2D
fig = plt.figure(figsize=(9.6, 9.4)); slots = [(0, 0, 2), (0, 2, 4), (1, 1, 3)]
gs = fig.add_gridspec(2, 4, left=0.07, right=0.98, bottom=0.06, top=0.96, wspace=0.55, hspace=0.32)
X, Y = np.meshgrid(np.linspace(0, 1, 60), np.linspace(0, 1, 60))
for k in range(3):
    r, c0, c1 = slots[k]; ax = fig.add_subplot(gs[r, c0:c1])
    Z = 60 + 15 * np.sin(2 * X + k) * np.cos(2 * Y)
    cf = ax.contourf(X, Y, Z, levels=np.arange(44, 80, 4), cmap="Greys")
    cs = ax.contour(X, Y, X + Y + 0.2 * k, levels=np.arange(0, 2.1, 0.25), colors="k", linewidths=0.9); ax.clabel(cs, fmt="%.2f", fontsize=9)
    ax.plot(rng.random(8), rng.random(8), "ko", ms=3.6, mfc="w", mew=0.7)
    ax.set_xlabel("u [-]", fontsize=10); ax.set_ylabel("v [-]", fontsize=10)
    ax.set_title(f"({'abc'[k]}) level {k + 1}", fontsize=12.5, fontweight="bold", loc="center")
cax = fig.add_axes([0.06, 0.255, 0.16, 0.028]); cb = fig.colorbar(cf, cax=cax, orientation="horizontal"); cb.set_label("z [%]", fontsize=12.5)
fig.legend(handles=[Line2D([], [], color="k", lw=1.2, label="w [-]"), Line2D([], [], ls="", marker="o", ms=5, mfc="w", mec="k", label="sample")],
           loc="center", bbox_to_anchor=(0.865, 0.27), fontsize=12.5, frameon=False)
fig.savefig(os.path.join(OUT, "fig2_panels.png"), dpi=160, bbox_inches="tight", pad_inches=0.06); plt.close(fig)
print("figs ok")
