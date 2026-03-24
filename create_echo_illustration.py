"""
牛（子牛）肺エコー プローブ当て位置 参考イラスト生成スクリプト
右側面・左側面の2図を生成し、PNGで保存する
"""

import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.patheffects as pe
import numpy as np
from matplotlib.patches import FancyArrowPatch, Arc, Wedge
from matplotlib.path import Path
import matplotlib.patches as mpatches

# フォント設定（日本語対応）
plt.rcParams['font.family'] = 'MS Gothic'
plt.rcParams['axes.unicode_minus'] = False

# =============================
# 描画ユーティリティ
# =============================

def draw_calf_body(ax, flip=False):
    """子牛の体幹シルエットを描画（右側面 or 左側面）"""
    sx = -1 if flip else 1

    # 胴体（楕円）
    body = mpatches.Ellipse((0, 0), width=5.2, height=2.8,
                             facecolor='#F5E6C8', edgecolor='#8B6914', linewidth=2.0, zorder=2)
    ax.add_patch(body)

    # 首・頭部
    neck_x = [sx * 2.0, sx * 2.6, sx * 3.1, sx * 3.3, sx * 2.9, sx * 2.2, sx * 2.0]
    neck_y = [0.8, 1.2, 1.4, 1.0, 0.5, 0.4, 0.8]
    ax.fill(neck_x, neck_y, color='#F5E6C8', zorder=3)
    ax.plot(neck_x, neck_y, color='#8B6914', linewidth=2.0, zorder=4)

    head = mpatches.Ellipse((sx * 3.4, 1.05), width=1.1, height=0.85,
                             facecolor='#F5E6C8', edgecolor='#8B6914', linewidth=2.0, zorder=4)
    ax.add_patch(head)

    # 鼻先（少し前に出る）
    nose = mpatches.Ellipse((sx * 3.95, 0.85), width=0.3, height=0.25,
                             facecolor='#E8C9A0', edgecolor='#8B6914', linewidth=1.5, zorder=5)
    ax.add_patch(nose)
    # 鼻孔
    ax.plot([sx * 3.88, sx * 3.88], [0.85, 0.85], 'o', color='#8B6914', markersize=3, zorder=6)
    ax.plot([sx * 3.95, sx * 3.95], [0.82, 0.82], 'o', color='#8B6914', markersize=3, zorder=6)

    # 目
    ax.plot([sx * 3.35], [1.25], 'o', color='#4A3000', markersize=5, zorder=6)

    # 耳
    ear_x = [sx * 3.1, sx * 3.0, sx * 3.3, sx * 3.3, sx * 3.1]
    ear_y = [1.4, 1.75, 1.65, 1.3, 1.4]
    ax.fill(ear_x, ear_y, color='#F5E6C8', zorder=3)
    ax.plot(ear_x, ear_y, color='#8B6914', linewidth=1.5, zorder=4)

    # 前肢
    fl_x = sx * 1.6
    # 上腕（前肢上部）
    ax.fill([fl_x - 0.22, fl_x + 0.22, fl_x + 0.22, fl_x - 0.22],
            [-1.0, -1.0, -0.3, -0.3], color='#F5E6C8', zorder=2)
    ax.plot([fl_x - 0.22, fl_x - 0.22, fl_x + 0.22, fl_x + 0.22, fl_x - 0.22],
            [-0.3, -1.0, -1.0, -0.3, -0.3], color='#8B6914', linewidth=1.5, zorder=3)
    # 前腕（肘から下）
    ax.fill([fl_x - 0.18, fl_x + 0.18, fl_x + 0.18, fl_x - 0.18],
            [-2.3, -2.3, -1.0, -1.0], color='#F5E6C8', zorder=2)
    ax.plot([fl_x - 0.18, fl_x - 0.18, fl_x + 0.18, fl_x + 0.18, fl_x - 0.18],
            [-1.0, -2.3, -2.3, -1.0, -1.0], color='#8B6914', linewidth=1.5, zorder=3)

    # 肘関節マーク
    elbow = mpatches.Circle((fl_x, -1.0), radius=0.12, facecolor='#FFD700',
                              edgecolor='#B8860B', linewidth=1.5, zorder=5)
    ax.add_patch(elbow)

    # 後肢
    rl_x = sx * (-1.6)
    ax.fill([rl_x - 0.22, rl_x + 0.22, rl_x + 0.22, rl_x - 0.22],
            [-1.0, -1.0, -0.3, -0.3], color='#F5E6C8', zorder=2)
    ax.plot([rl_x - 0.22, rl_x - 0.22, rl_x + 0.22, rl_x + 0.22, rl_x - 0.22],
            [-0.3, -1.0, -1.0, -0.3, -0.3], color='#8B6914', linewidth=1.5, zorder=3)
    ax.fill([rl_x - 0.18, rl_x + 0.18, rl_x + 0.18, rl_x - 0.18],
            [-2.3, -2.3, -1.0, -1.0], color='#F5E6C8', zorder=2)
    ax.plot([rl_x - 0.18, rl_x - 0.18, rl_x + 0.18, rl_x + 0.18, rl_x - 0.18],
            [-1.0, -2.3, -2.3, -1.0, -1.0], color='#8B6914', linewidth=1.5, zorder=3)

    return fl_x


def draw_ribs_and_scanzone(ax, flip=False):
    """肋骨・肋間番号・スキャンゾーン・プローブを描画"""
    sx = -1 if flip else 1

    # 肋骨の起点（背側、体幹上部）から腹側に向けて描く
    # 右側面: 前から第1肋骨は前肢の直後付近。第2〜6肋間を描く
    # 肋骨の x 座標（前肢からの距離）
    # 前肢中心は sx * 1.6 付近、肋骨間隔は約0.35

    rib_start_x = sx * 1.55  # 前肢のやや後ろ（第1肋骨）
    rib_spacing = sx * (-0.35)  # 後方へ
    n_ribs = 8  # 描画するリブ数（1〜8）

    rib_xs = []
    for i in range(n_ribs):
        rx = rib_start_x + rib_spacing * i
        rib_xs.append(rx)

    # 肺野（胸郭内）の背景
    # 胸郭は第1肋骨〜最後肋骨、背側〜腹側
    lung_top_y = 0.9
    lung_bot_y = -0.85

    # スキャンゾーン（第2〜第6肋間）をハイライト
    # 肋間 n = n番目と(n+1)番目の肋骨の間
    scan_ics = [2, 3, 4, 5, 6]  # スキャンする肋間番号

    scan_colors = {
        2: '#FF6B6B',
        3: '#FF9F43',
        4: '#54A0FF',
        5: '#5F27CD',
        6: '#00D2D3',
    }

    for ic in scan_ics:
        idx = ic - 1  # 肋間icはrib_xs[ic-1]とrib_xs[ic]の間
        if idx + 1 < len(rib_xs):
            x1 = rib_xs[idx]
            x2 = rib_xs[idx + 1]
            xmin = min(x1, x2) + 0.01
            xmax = max(x1, x2) - 0.01
            rect = mpatches.FancyBboxPatch(
                (xmin, lung_bot_y - 0.05),
                abs(xmax - xmin),
                (lung_top_y - lung_bot_y + 0.1),
                boxstyle="round,pad=0.01",
                facecolor=scan_colors[ic], alpha=0.25,
                edgecolor=scan_colors[ic], linewidth=0, zorder=3
            )
            ax.add_patch(rect)

    # 肋骨ラインを描画（縦線）
    for i, rx in enumerate(rib_xs):
        rib_num = i + 1
        if rib_num <= 7:
            ax.plot([rx, rx], [lung_bot_y - 0.1, lung_top_y + 0.05],
                    color='#C0A070', linewidth=1.8, zorder=4, alpha=0.7)

    # 肋間番号ラベル
    for ic in scan_ics:
        idx = ic - 1
        if idx + 1 < len(rib_xs):
            x1 = rib_xs[idx]
            x2 = rib_xs[idx + 1]
            mx = (x1 + x2) / 2
            # 肋間番号（背側・上方）
            ax.text(mx, lung_top_y + 0.22, f'第{ic}肋間',
                    ha='center', va='bottom', fontsize=7, color='#333333',
                    fontweight='bold', zorder=6, rotation=0)

    # プローブ位置マーカーをスキャンゾーンの中央付近に描画
    probe_y = 0.05  # 体表面付近（腹側〜中央）
    for ic in scan_ics:
        idx = ic - 1
        if idx + 1 < len(rib_xs):
            x1 = rib_xs[idx]
            x2 = rib_xs[idx + 1]
            mx = (x1 + x2) / 2
            c = scan_colors[ic]
            # プローブ（台形型）
            pw = 0.08
            ph = 0.18
            probe = mpatches.FancyBboxPatch(
                (mx - pw, probe_y - ph / 2), pw * 2, ph,
                boxstyle="round,pad=0.01",
                facecolor=c, edgecolor='white', linewidth=1.5, zorder=7, alpha=0.9
            )
            ax.add_patch(probe)
            # 矢印（背側→腹側のスキャン方向）
            ax.annotate('', xy=(mx, probe_y - ph / 2 - 0.15),
                        xytext=(mx, probe_y + ph / 2 + 0.12),
                        arrowprops=dict(arrowstyle='->', color=c, lw=2), zorder=8)

    return rib_xs, scan_colors


def draw_legend(ax, scan_colors, side_label):
    """凡例を描画"""
    legend_x = 3.5
    legend_y = 1.2
    ax.text(legend_x, legend_y, f'【{side_label}肺 スキャン肋間】',
            ha='center', va='top', fontsize=8, fontweight='bold', color='#333333', zorder=10)

    scan_ics = [2, 3, 4, 5, 6]
    short = 'R' if side_label == '右' else 'L'
    for j, ic in enumerate(scan_ics):
        ly = legend_y - 0.28 - j * 0.24
        c = scan_colors[ic]
        sq = mpatches.FancyBboxPatch((legend_x - 0.7, ly - 0.09), 0.2, 0.18,
                                      boxstyle="round,pad=0.01",
                                      facecolor=c, edgecolor='white', linewidth=1, zorder=10)
        ax.add_patch(sq)
        ax.text(legend_x - 0.4, ly, f'{short}{ic}  第{ic}肋間',
                ha='left', va='center', fontsize=7.5, color='#222222', zorder=10)


# =============================
# メイン描画
# =============================

fig, axes = plt.subplots(1, 2, figsize=(16, 7))
fig.patch.set_facecolor('#FAFAF7')

titles = ['右側面 (Right side)', '左側面 (Left side)']
flips = [False, True]
side_labels = ['右', '左']

for ax, flip, title, side_label in zip(axes, flips, titles, side_labels):
    ax.set_facecolor('#EEF4FF')
    ax.set_xlim(-5, 5)
    ax.set_ylim(-3.0, 2.8)
    ax.set_aspect('equal')
    ax.axis('off')

    # 背景グリッドなし、タイトル
    ax.set_title(title, fontsize=11, fontweight='bold', color='#1A1A2E', pad=8)

    # 胴体描画
    fl_x = draw_calf_body(ax, flip=flip)

    # 肋骨・スキャンゾーン描画
    rib_xs, scan_colors = draw_ribs_and_scanzone(ax, flip=flip)

    # 肘関節ラベル
    sx = -1 if flip else 1
    ax.annotate('肘関節\n(基準点)', xy=(fl_x, -1.0), xytext=(fl_x + sx * (-0.6), -1.55),
                fontsize=7.5, ha='center', color='#8B6914', fontweight='bold',
                arrowprops=dict(arrowstyle='->', color='#B8860B', lw=1.5),
                bbox=dict(boxstyle='round,pad=0.2', facecolor='#FFF9E6', edgecolor='#B8860B', alpha=0.9),
                zorder=9)

    # スキャン方向注釈
    ax.text(0, -2.75, '▲ 背側から腹側に向けてプローブをスライド。各肋間でエコー画像を確認',
            ha='center', va='bottom', fontsize=7.5, color='#555555',
            style='italic', zorder=10)

    # 凡例
    lx = (3.5 if not flip else -3.5)
    legend_y = 1.3
    ax.text(lx, legend_y, f'【{side_label}肺  スキャン肋間】',
            ha='center', va='top', fontsize=8, fontweight='bold', color='#333333', zorder=10)
    scan_ics = [2, 3, 4, 5, 6]
    short = 'R' if side_label == '右' else 'L'
    for j, ic in enumerate(scan_ics):
        ly = legend_y - 0.3 - j * 0.24
        c = scan_colors[ic]
        sq_x = lx - 0.65
        sq = mpatches.FancyBboxPatch((sq_x, ly - 0.09), 0.22, 0.18,
                                      boxstyle="round,pad=0.01",
                                      facecolor=c, edgecolor='white', linewidth=1, zorder=10)
        ax.add_patch(sq)
        ax.text(sq_x + 0.32, ly, f'{short}{ic}  第{ic}肋間',
                ha='left', va='center', fontsize=7.5, color='#222222', zorder=10)

# タイトルと注釈
fig.suptitle('子牛 肺エコー検診  プローブ当て位置ガイド',
             fontsize=14, fontweight='bold', color='#1A1A2E', y=0.98)

# 下部共通注釈
fig.text(0.5, 0.01,
         '※ 体表を清潔にし、ゲルを塗布してから検査。前肢後方の肋間に沿って、背側→腹側方向にスキャン。'
         'R2〜R6（右肺）、L2〜L6（左肺）を標準スキャン範囲とする。',
         ha='center', fontsize=8, color='#555555', style='italic')

plt.tight_layout(rect=[0, 0.04, 1, 0.96])

out_path = r'C:\Users\user\OneDrive\デスクトップ\肺エコープロジェクト\肺エコー_プローブ位置ガイド.png'
plt.savefig(out_path, dpi=150, bbox_inches='tight', facecolor=fig.get_facecolor())
print(f'保存完了: {out_path}')
plt.close()
