# -*- coding: utf-8 -*-
"""
slide_image_export.py
─────────────────────────────────────────────────────────────────────
スライド（PowerPoint / Word）にそのまま貼れる PNG 画像を生成するモジュール。

提供する2つの画像：
  ① make_demand_chart_png()  …… 月別「使用量(積み上げ：使用量+削減量) ＋ デマンド折れ線」グラフ
  ② make_control_list_png()  …… 空調の制御可否リスト（2カラム、最大100件程度を1枚に）

いずれも PNG の bytes を返す（Streamlit の download_button にそのまま渡せる）。
out_path を渡すとファイルにも保存する。

依存: matplotlib, pillow（日本語フォントは Windows 標準の Meiryo / Yu Gothic を自動選択）
"""

from __future__ import annotations

import io
import os
import logging
from typing import Sequence



import matplotlib
matplotlib.use("Agg")  # GUI 不要・サーバー描画
import matplotlib.pyplot as plt
import matplotlib.image as mpimg
import matplotlib.patheffects as pe
from matplotlib import font_manager as fm
from matplotlib.patches import FancyBboxPatch, Rectangle, Polygon, Ellipse

_MODULE_DIR = os.path.dirname(os.path.abspath(__file__))

# ─────────────────────────────────────────────────────────────────────
# 配色（参考資料に準拠）
# ─────────────────────────────────────────────────────────────────────
NAVY   = "#13315C"   # 使用量バー / バナー
GREEN  = "#3DAE4E"   # 削減量バー / 削減コールアウト
ORANGE = "#E8730C"   # デマンド折れ線 / ピーク強調
GRID   = "#D9DDE3"
OK_BG  = "#7CD17F"   # 制御可（○）セル
NG_BG  = "#DCDCDC"   # 制御不可（×）セル
HEAD_BG = NAVY


def set_jp_font() -> str:
    """利用可能な日本語フォントを matplotlib に設定し、フォント名を返す。"""
    fm._load_fontmanager(try_read_cache=False)
    names = {f.name for f in fm.fontManager.ttflist}
   print(names)

    for cand in (
        "Noto Sans CJK JP",
        "Noto Sans JP",
        "IPAexGothic",
        "IPAGothic",
        "Meiryo",
        "Yu Gothic",
        "BIZ UDGothic",
        "MS Gothic",
        "MS PGothic",
    ):
        if cand in names:
            plt.rcParams["font.family"] = cand
            plt.rcParams["axes.unicode_minus"] = False
            return cand

    plt.rcParams["font.family"] = "sans-serif"
    plt.rcParams["axes.unicode_minus"] = False
    return "sans-serif"


def _short_month_labels(months: Sequence) -> list[str]:
    """x軸ラベルを「○月」に整形する。先頭の月、および年が変わる月（＝翌年の1月など）には
    "2026年\n1月" のように年を併記する。

    日時(datetime/date/Timestamp)・"2025-05-01 00:00:00"・"2025/05"・"5月"
    などの混在を許容。年が拾えない（"5月"等）場合は年の変わり目を判定できないためそのまま「○月」。
    """
    import re
    out: list[str] = []
    prev_year = None
    for m in months:
        year = mon = None
        if hasattr(m, "year") and hasattr(m, "month"):      # datetime / date / Timestamp
            year, mon = int(m.year), int(m.month)
        else:
            s = str(m).strip()
            mt = re.search(r"(\d{4})\s*[/\-\.年]\s*(\d{1,2})", s)   # 2025-05 / 2025/5 / 2025年5
            if mt:
                year, mon = int(mt.group(1)), int(mt.group(2))
            else:
                mt2 = re.search(r"(\d{1,2})", s)                    # "5月" / "5"
                if mt2:
                    mon = int(mt2.group(1))
        if mon is None:
            out.append(str(m))                              # 解釈不能はそのまま
            continue
        if year is not None and year != prev_year:          # 先頭 or 年が変わった月に年を併記
            out.append(f"{year}年\n{mon}月")
            prev_year = year
        else:
            out.append(f"{mon}月")
    return out


def _fig_to_png(fig, out_path: str | None, dpi: int) -> bytes:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=dpi, bbox_inches="tight",
                facecolor="white", pad_inches=0.15)
    plt.close(fig)
    data = buf.getvalue()
    if out_path:
        with open(out_path, "wb") as f:
            f.write(data)
    return data


# ═════════════════════════════════════════════════════════════════════
# ① 使用量(積み上げ) ＋ デマンド折れ線
# ═════════════════════════════════════════════════════════════════════
def make_demand_chart_png(
    months: Sequence[str],
    usage: Sequence[float],          # 使用量(kWh) … 制御後に残る分（バー下段）
    reduction: Sequence[float],      # 削減量(kWh) … 上に積む分（バー上段・緑）
    demand: Sequence[float],         # 最大デマンド(kW) … 折れ線
    *,
    target_units: int | None = None,         # 制御対象台数（フッターに表示）
    title: str | None = None,
    out_path: str | None = None,
    dpi: int = 200,
) -> bytes:
    """月別の使用量(積み上げ) ＋ デマンド折れ線グラフを PNG 化して返す。

    バー＝使用量。下段(濃紺)=制御後に残る使用量、上段(緑)=DPSで削減できる量。
    バー全体の高さ＝現状の使用量。折れ線(橙)＝最大デマンド(右軸)。
    """
    n = len(months)
    usage = [float(x) for x in usage]
    reduction = [float(x) for x in reduction]
    demand = [float(x) for x in demand]
    total_usage = sum(usage)          # 棒(濃紺)＝実使用量。緑は視覚化のため上積みするだけ
    total_reduc = sum(reduction)
    peak = max(demand)
    peak_i = demand.index(peak)

    set_jp_font()
    fig, axL = plt.subplots(figsize=(11.0, 5.6))
    axR = axL.twinx()

    x = range(n)
    bw = 0.64

    # ── 積み上げ棒（下：使用量＝濃紺／上：削減量＝緑） ──
    # 削減量は実数では細すぎて見えないため、見栄え優先で表示高さを拡大（ラベルは実数）
    max_u = max(usage) if usage else 0.0
    max_r = max(reduction) if reduction else 0.0
    disp_scale = ((0.13 * max_u) / max_r) if max_r > 0 else 1.0
    disp_red = [r * disp_scale for r in reduction]

    # 両軸の上限（ラベル衝突回避の座標換算に使う）。全て0でもゼロ割にならないよう下限1.0
    _ml = max((u + r for u, r in zip(usage, disp_red)), default=0.0)
    ymax_L = (_ml * 1.22) if _ml > 0 else 1.0
    ymax_R = peak * 1.18 if peak > 0 else 1.0

    axL.bar(x, usage, bw, color=NAVY, zorder=3, label="使用量(kWh)")
    axL.bar(x, disp_red, bw, bottom=usage, color=GREEN, zorder=3, label="削減量(kWh)")

    # データラベル（縁取り文字：文字の淵にストロークを付けて視認性を確保）
    stroke_u = [pe.withStroke(linewidth=2.4, foreground="#0A1B33")]   # 白文字＋濃紺の淵
    stroke_r = [pe.withStroke(linewidth=2.4, foreground="#15531F")]   # 白文字＋濃緑の淵
    stroke_d = [pe.withStroke(linewidth=2.6, foreground="white")]     # 橙文字＋白の淵
    for i in range(n):
        if usage[i] > 0:
            axL.text(i, usage[i] / 2, f"{usage[i]:,.0f}", ha="center", va="center",
                     color="white", fontsize=8, fontweight="bold",
                     path_effects=stroke_u, zorder=4)
        if reduction[i] > 0:
            axL.text(i, usage[i] + disp_red[i] / 2, f"{reduction[i]:,.0f}",
                     ha="center", va="center", color="white", fontsize=8,
                     fontweight="bold", path_effects=stroke_r, zorder=4)

    # ── デマンド折れ線（右軸：マーカー無し） ──
    axR.plot(x, demand, color=ORANGE, lw=2.8, zorder=5)
    # 数値ラベルは「線の点」と「棒の頂上」の高い方の上に置き、棒上の緑ラベルとの重なりを自動回避
    for i in range(n):
        bartop_R = (usage[i] + disp_red[i]) / ymax_L * ymax_R   # 棒頂上を右軸スケールへ換算
        y_anchor = max(demand[i], bartop_R)
        axR.annotate(f"{demand[i]:,.0f}", (i, y_anchor),
                     textcoords="offset points", xytext=(0, 12),
                     ha="center", fontsize=8, color=ORANGE, fontweight="bold",
                     path_effects=stroke_d, zorder=6, clip_on=False)

    # ── 軸の体裁 ──
    axL.set_xticks(list(x))
    axL.set_xticklabels(_short_month_labels(months))
    axL.set_ylabel("Usage (kWh)", fontsize=10)
    axR.set_ylabel("Max Demand (kW)", fontsize=10)
    axL.set_ylim(0, ymax_L)
    axR.set_ylim(0, ymax_R)
    axL.set_axisbelow(True)
    axL.yaxis.grid(True, color=GRID, lw=0.8)
    axL.set_xlim(-0.7, n - 0.3)
    for sp in ("top",):
        axL.spines[sp].set_visible(False)
        axR.spines[sp].set_visible(False)
    axL.tick_params(labelsize=9)
    axR.tick_params(labelsize=9)
    axL.yaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda v, _: f"{v:,.0f}"))

    if title:
        axL.set_title(title, fontsize=13, fontweight="bold", color=NAVY, pad=10)

    # ── フッターバナー（年間合計） ──
    banner = (f"年間総使用電力量： {total_usage:,.0f} kWh　/　"
              f"年間最大需要電力： {peak:,.0f} kW")
    fig.text(0.5, -0.02, banner, ha="center", va="top", fontsize=11.5,
             color="white", fontweight="bold",
             bbox=dict(boxstyle="round,pad=0.45", fc=NAVY, ec="none"))

    # ── フッター2行目：制御対象台数 ＋ 年間削減量（緑の囲み）を横並び ──
    units_text = f"対象設備：制御対象台数 室外機 {target_units} 台" if target_units is not None else ""
    reduc_text = f"年間約 {total_reduc:,.0f} kWh 削減" if total_reduc > 0 else ""
    if units_text and reduc_text:
        fig.text(0.49, -0.105, units_text, ha="right", va="top",
                 fontsize=11, color="#222222", fontweight="bold")
        fig.text(0.52, -0.105, reduc_text, ha="left", va="top",
                 fontsize=11, color="white", fontweight="bold",
                 bbox=dict(boxstyle="round,pad=0.35", fc=GREEN, ec="none"))
    elif reduc_text:
        fig.text(0.5, -0.105, reduc_text, ha="center", va="top",
                 fontsize=11, color="white", fontweight="bold",
                 bbox=dict(boxstyle="round,pad=0.35", fc=GREEN, ec="none"))
    elif units_text:
        fig.text(0.5, -0.105, units_text, ha="center", va="top",
                 fontsize=11, color="#222222", fontweight="bold")

    return _fig_to_png(fig, out_path, dpi)


# ═════════════════════════════════════════════════════════════════════
# ② 制御可否リスト（2カラム）
# ═════════════════════════════════════════════════════════════════════
# items: 各行 dict {"系統名","設置場所","メーカー","型式","制御可否"(bool or "○"/"×")}
_COLS = ["系統名", "設置場所", "メーカー", "型式", "制御可否"]
_COL_W = [0.30, 0.20, 0.16, 0.26, 0.08]   # 各カラム相対幅（合計1.0）
# 列ごとの基準相対幅（列指定で一部を除外した場合はこの比で再配分）
_COL_W_MAP = dict(zip(_COLS, _COL_W))


def _resolve_cols(cols):
    """使用する列リストを正規化。None なら全列。順序は _COLS 基準。"""
    if not cols:
        return list(_COLS)
    keep = [c for c in _COLS if c in cols]
    return keep or list(_COLS)


def _norm_ok(v) -> bool:
    if isinstance(v, bool):
        return v
    return str(v).strip() in ("○", "◯", "〇", "o", "O", "1", "True", "true", "可")


def _draw_table(ax, items, x0, x1, cols=None):
    """ax の [x0, x1]×[0,1] 領域に1カラム分の表を描く。cols で列を選択可。"""
    cols = _resolve_cols(cols)
    nrows = len(items) + 1  # ヘッダ含む
    top, bottom = 0.97, 0.02
    rh = (top - bottom) / nrows
    w = x1 - x0
    # 選択列の相対幅を正規化して x 境界を作る
    ws = [_COL_W_MAP.get(c, 0.2) for c in cols]
    tot = sum(ws) or 1.0
    ws = [v / tot for v in ws]
    xs = [x0]
    for cw in ws:
        xs.append(xs[-1] + cw * w)

    fs = max(5.5, min(8.0, 8.0 * (26 / max(nrows, 26))))  # 行数に応じ縮小

    # ヘッダ
    y = top - rh
    for ci, col in enumerate(cols):
        ax.add_patch(Rectangle((xs[ci], y), xs[ci + 1] - xs[ci], rh,
                               facecolor=HEAD_BG, edgecolor="white", lw=0.6, zorder=2))
        ax.text((xs[ci] + xs[ci + 1]) / 2, y + rh / 2, col, ha="center", va="center",
                color="white", fontsize=fs, fontweight="bold", zorder=3)
    # データ行
    for ri, it in enumerate(items):
        y = top - rh * (ri + 2)
        ok = _norm_ok(it.get("制御可否"))
        for ci, key in enumerate(cols):
            cx0, cx1 = xs[ci], xs[ci + 1]
            if key == "制御可否":
                ax.add_patch(Rectangle((cx0, y), cx1 - cx0, rh,
                                       facecolor=OK_BG if ok else NG_BG,
                                       edgecolor="white", lw=0.6, zorder=2))
                ax.text((cx0 + cx1) / 2, y + rh / 2, "○" if ok else "×",
                        ha="center", va="center", fontsize=fs + 1,
                        color="#0B3D14" if ok else "#666666", zorder=3)
            else:
                ax.add_patch(Rectangle((cx0, y), cx1 - cx0, rh,
                                       facecolor="white", edgecolor="#C9CDD3",
                                       lw=0.6, zorder=2))
                txt = str(it.get(key, ""))
                ha = "center" if key in ("設置場所", "メーカー") else "left"
                tx = (cx0 + cx1) / 2 if ha == "center" else cx0 + 0.004
                ax.text(tx, y + rh / 2, txt, ha=ha, va="center",
                        fontsize=fs, color="#222222", zorder=3)


# 1ページ（=1スライド）あたりの最大件数。2カラム(25行×2)を基準に16:9へ収める
MAX_PER_PAGE = 50


def make_control_list_png(
    items: Sequence[dict],
    *,
    total_units: int | None = None,        # 拠点室外機 総数
    controllable_units: int | None = None, # 制御可能 台数
    title: str = "制御可否リスト",
    out_path: str | None = None,
    dpi: int = 200,
    page: int | None = None,               # ページ番号（分割時）
    total_pages: int | None = None,        # 総ページ数（分割時）
    fig_h: float | None = None,            # 図の高さ(inch)を明示指定（列幅=7.1"固定のまま行だけ収める）
    cols: Sequence[str] | None = None,     # 出力する列（None=全列）。系統名/設置場所/メーカー/型式/制御可否
) -> bytes:
    """制御可否リスト1ページ分を2カラムの表として PNG 化して返す（スライド枠幅に合わせて貼付）。

    items を左右2カラムに分割。図の横幅は固定(13")で、貼付側で枠幅に合わせて拡縮する。
    fig_h を渡すと高さを明示指定できる。複数ページは make_control_list_pngs() を使う。
    """
    items = list(items)
    if controllable_units is None:
        controllable_units = sum(1 for it in items if _norm_ok(it.get("制御可否")))
    if total_units is None:
        total_units = len(items)

    # 左右に分割（左を多めに）
    half = (len(items) + 1) // 2
    left, right = items[:half], items[half:]

    set_jp_font()
    max_rows = max(len(left), len(right), 1)
    if fig_h is None:                            # 未指定なら行数から自動決定
        fig_h = max(3.2, 1.1 + max_rows * 0.255)
    fig, ax = plt.subplots(figsize=(13.0, fig_h))
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")

    # タイトル＋台数（台数はタイトル右端のすぐ右に配置）
    page_tag = f"（{page}/{total_pages}）" if page and total_pages and total_pages > 1 else ""
    head = f"■ {title}{page_tag}"
    sub = f"拠点室外機 全{total_units}台　制御可能室外機 {controllable_units}台"
    head_x = 0.012
    t_head = fig.text(head_x, 0.985, head, ha="left", va="top", fontsize=14,
                      fontweight="bold", color=NAVY)
    sub_x = head_x + 0.16
    try:
        fig.canvas.draw()
        bb = t_head.get_window_extent(renderer=fig.canvas.get_renderer())
        sub_x = fig.transFigure.inverted().transform((bb.x1, bb.y0))[0] + 0.015
    except Exception:
        pass
    fig.text(sub_x, 0.978, sub, ha="left", va="top", fontsize=11.5, color="#222222")

    ax2 = fig.add_axes([0.01, 0.02, 0.98, 0.90])
    ax2.set_xlim(0, 1)
    ax2.set_ylim(0, 1)
    ax2.axis("off")
    _draw_table(ax2, left, 0.0, 0.485, cols)
    if right:
        _draw_table(ax2, right, 0.515, 1.0, cols)

    return _fig_to_png(fig, out_path, dpi)


def make_control_list_pngs(
    items: Sequence[dict],
    *,
    total_units: int | None = None,
    controllable_units: int | None = None,
    title: str = "制御可否リスト",
    max_per_page: int = MAX_PER_PAGE,
    out_path_prefix: str | None = None,    # 例 "list" → list_1.png, list_2.png
    dpi: int = 200,
    fig_h: float | None = None,            # 各ページ図の高さ(inch)を明示指定（列幅一定のまま）
    cols: Sequence[str] | None = None,     # 出力する列（None=全列）
) -> list[bytes]:
    """制御可否リストを、1スライドに収まる件数で自動的に複数ページへ分割して PNG 化する。

    件数が max_per_page 以下なら1枚、超えたら必要枚数に分割。台数は全体合計を各ページに表示し、
    タイトル右に「(1/2)」等のページ表記を付ける。戻り値は各ページの PNG bytes のリスト。
    """
    items = list(items)
    if total_units is None:
        total_units = len(items)
    if controllable_units is None:
        controllable_units = sum(1 for it in items if _norm_ok(it.get("制御可否")))

    pages = [items[i:i + max_per_page] for i in range(0, len(items), max_per_page)] or [[]]
    n_pages = len(pages)
    out = []
    for k, page_items in enumerate(pages, start=1):
        op = f"{out_path_prefix}_{k}.png" if out_path_prefix else None
        out.append(make_control_list_png(
            page_items, total_units=total_units, controllable_units=controllable_units,
            title=title, out_path=op, dpi=dpi, page=k, total_pages=n_pages, fig_h=fig_h,
            cols=cols))
    return out


# ═════════════════════════════════════════════════════════════════════
# ③ サマリKPIカード（横並びカード）
# ═════════════════════════════════════════════════════════════════════
# cards: 各カード dict
#   {"icon": "money.png"(モジュール直下) or 絶対パス, "title": str,
#    "value": str, "subtitle": str, "color": 値の色(省略時 GREEN)}
def _resolve_icon(path: str) -> str | None:
    if not path:
        return None
    if os.path.isabs(path) and os.path.exists(path):
        return path
    cand = os.path.join(_MODULE_DIR, path)
    return cand if os.path.exists(cand) else (path if os.path.exists(path) else None)


def _fit_axtext(fig, ax, x, y, s, *, color, base_fs, max_w_in, weight="bold"):
    """ax(インチ座標)にテキストを置き、max_w_in を超える幅なら自動縮小。"""
    t = ax.text(x, y, s, ha="center", va="center", fontsize=base_fs,
                fontweight=weight, color=color)
    try:
        fig.canvas.draw()
        bb = t.get_window_extent(fig.canvas.get_renderer())
        w_in = bb.width / fig.dpi
        if w_in > max_w_in and w_in > 0:
            t.set_fontsize(base_fs * max_w_in / w_in)
    except Exception:
        pass
    return t


# 値の末尾にある単位（円・年・台・%等）を分離する。数字部分は大きく、単位は小さく描く用。
_UNIT_SUFFIXES = ("ヶ月", "万円", "円", "年", "台", "kW", "%")


def _split_unit(s: str):
    s = str(s)
    for u in _UNIT_SUFFIXES:
        if s.endswith(u) and len(s) > len(u):
            return s[:-len(u)], u
    return s, ""


def _fit_value_unit(fig, ax, cx, y, main, unit, *, color, base_fs, max_w_in,
                    unit_ratio=0.52):
    """数字(main)を大きく、単位(unit)を小さく、全体を cx 中心に配置。max_w_in で自動縮小。"""
    unit_fs = base_fs * unit_ratio
    gap = 0.03
    tm = ax.text(0, y, main, ha="left", va="center", fontsize=base_fs,
                 fontweight="bold", color=color, zorder=4)
    tu = ax.text(0, y, unit, ha="left", va="center", fontsize=unit_fs,
                 fontweight="bold", color=color, zorder=4)
    try:
        fig.canvas.draw()
        r = fig.canvas.get_renderer()
        mw = tm.get_window_extent(r).width / fig.dpi
        uw = tu.get_window_extent(r).width / fig.dpi
        total = mw + gap + uw
        if total > max_w_in and total > 0:
            sc = max_w_in / total
            base_fs *= sc; unit_fs *= sc
            tm.set_fontsize(base_fs); tu.set_fontsize(unit_fs)
            fig.canvas.draw()
            mw = tm.get_window_extent(r).width / fig.dpi
            uw = tu.get_window_extent(r).width / fig.dpi
            total = mw + gap + uw
        left = cx - total / 2.0
        tm.set_position((left, y))
        tu.set_position((left + mw + gap, y))
    except Exception:
        tm.set_position((cx, y)); tm.set_ha("center")
        tu.set_position((cx, y)); tu.set_alpha(0)
    return tm, tu


def make_summary_cards_png(
    cards: Sequence[dict],
    *,
    out_path: str | None = None,
    dpi: int = 200,
    fig_w: float = 13.0,
    fig_h: float = 4.7,
) -> bytes:
    """サマリKPIを横並びカード（白角丸・影付き）として PNG 化して返す。

    各カード：上にアイコン、下に タイトル / 大きな数値 / 補足（括弧書き）。
    数値はカード幅に合わせて自動縮小。アイコンはモジュール同階層のファイル名で指定可。
    """
    cards = list(cards)
    n = max(len(cards), 1)
    set_jp_font()
    fig = plt.figure(figsize=(fig_w, fig_h))

    outer, gap = 0.018, 0.022
    avail = 1 - 2 * outer - (n - 1) * gap
    cw = avail / n
    y0, y1 = 0.05, 0.95
    ch = y1 - y0

    SUBT = "#6B7280"
    for i, card in enumerate(cards):
        x0 = outer + i * (cw + gap)
        pw, ph = cw * fig_w, ch * fig_h          # カード物理サイズ(インチ)
        axc = fig.add_axes([x0, y0, cw, ch])
        axc.set_xlim(0, pw)
        axc.set_ylim(0, ph)
        axc.axis("off")

        # 影 → 白カード（角丸はインチ座標なので均等）
        axc.add_patch(FancyBboxPatch(
            (0.10, 0.02), pw - 0.18, ph - 0.14,
            boxstyle="round,pad=0,rounding_size=0.22",
            fc="#00000018", ec="none", zorder=1))
        axc.add_patch(FancyBboxPatch(
            (0.07, 0.10), pw - 0.18, ph - 0.16,
            boxstyle="round,pad=0,rounding_size=0.22",
            fc="white", ec="#E5E7EB", lw=1.2, zorder=2))

        # アイコン（白カード上部中央）
        icon = _resolve_icon(str(card.get("icon", "")))
        if icon:
            try:
                img = mpimg.imread(icon)
                isz = min(0.95, ph * 0.30)        # インチ
                w_frac, h_frac = isz / fig_w, isz / fig_h
                cx_frac = x0 + (pw / 2) / fig_w
                cy_frac = y0 + (ph * 0.72) / fig_h
                axi = fig.add_axes([cx_frac - w_frac / 2, cy_frac - h_frac / 2,
                                    w_frac, h_frac], zorder=3)
                axi.imshow(img)
                axi.axis("off")
            except Exception:
                pass

        color = card.get("color", GREEN)
        # タイトル / 値 / 補足
        _fit_axtext(fig, axc, pw / 2, ph * 0.46, str(card.get("title", "")),
                    color=NAVY, base_fs=17, max_w_in=pw * 0.86)
        axc.texts[-1].set_zorder(4)
        _main, _unit = _split_unit(card.get("value", ""))
        if _unit:
            _fit_value_unit(fig, axc, pw / 2, ph * 0.31, _main, _unit,
                            color=color, base_fs=32, max_w_in=pw * 0.86)
        else:
            _fit_axtext(fig, axc, pw / 2, ph * 0.31, str(card.get("value", "")),
                        color=color, base_fs=32, max_w_in=pw * 0.86)
            axc.texts[-1].set_zorder(4)
        sub = str(card.get("subtitle", ""))
        if sub:
            _fit_axtext(fig, axc, pw / 2, ph * 0.155, sub,
                        color=SUBT, base_fs=10.5, max_w_in=pw * 0.90, weight="normal")
            axc.texts[-1].set_zorder(4)

    return _fig_to_png(fig, out_path, dpi)


# ═════════════════════════════════════════════════════════════════════
# ④ 制御%別 使用量バー（10年収支スライド 左）
# ═════════════════════════════════════════════════════════════════════
def make_usage_bars_png(
    categories: Sequence[str],
    values: Sequence[float],
    *,
    title: str = "空調の年間電力使用量（万kWh/年）─ デマンド制御% 別",
    unit: str = "万",
    highlight_first: bool = True,
    out_path: str | None = None,
    dpi: int = 200,
    fig_w: float = 6.2,
    fig_h: float = 3.5,
) -> bytes:
    """制御強度別の年間使用量バー。先頭（現状）は濃紺、以降は緑。棒上に「20.0万」ラベル。"""
    set_jp_font()
    cats = list(categories)
    vals = [float(v) for v in values]
    n = len(vals)
    xs = list(range(n))
    colors = [NAVY if (i == 0 and highlight_first) else GREEN for i in range(n)]

    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    ax.bar(xs, vals, width=0.62, color=colors, zorder=3)

    vmax = max(vals) if vals else 1.0
    for x, v in zip(xs, vals):
        ax.text(x, v + vmax * 0.03, f"{v:.1f}{unit}", ha="center", va="bottom",
                fontsize=12, fontweight="bold", color="#222222", zorder=4)

    ax.set_xticks(xs)
    ax.set_xticklabels(cats, fontsize=11, color="#333333")
    ax.set_ylim(0, vmax * 1.20)
    ax.margins(x=0.04)
    ax.tick_params(axis="y", labelsize=9, color="#B9C0CA")
    ax.tick_params(axis="x", length=0)
    for sp in ("top", "right"):
        ax.spines[sp].set_visible(False)
    ax.spines["left"].set_color("#B9C0CA")
    ax.spines["bottom"].set_color("#B9C0CA")
    ax.yaxis.grid(True, color=GRID, linewidth=0.8, zorder=0)
    ax.set_axisbelow(True)
    ax.set_title(title, fontsize=11.5, fontweight="bold", color=NAVY, loc="left", pad=10)
    return _fig_to_png(fig, out_path, dpi)


# ═════════════════════════════════════════════════════════════════════
# ⑤ 累積収支の折れ線（10年収支スライド 右）
# ═════════════════════════════════════════════════════════════════════
def make_cumulative_balance_png(
    year_labels: Sequence[str],
    values: Sequence[float],
    *,
    payback_years: float | None = None,
    title: str = "累積収支（初期費用控除後・万円）",
    out_path: str | None = None,
    dpi: int = 200,
    fig_w: float = 6.2,
    fig_h: float = 3.5,
) -> bytes:
    """累積収支の折れ線（緑・ダイヤマーカー）。0ラインを破線、終点に「+824」、回収点に吹き出し。"""
    set_jp_font()
    labels = list(year_labels)
    vals = [float(v) for v in values]
    n = len(vals)
    xs = list(range(n))
    LINE = "#2E7D32"

    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    ax.axhline(0, color="#9AA3AE", linewidth=1.0, linestyle=(0, (5, 4)), zorder=1)
    ax.plot(xs, vals, color=LINE, linewidth=2.6, marker="D", markersize=6,
            markerfacecolor=LINE, markeredgecolor="white", markeredgewidth=0.8, zorder=3)

    y_lo, y_hi = min(vals), max(vals)
    yr = (y_hi - y_lo) or 1.0
    last = vals[-1]
    ax.text(xs[-1], last + yr * 0.05, f"{'+' if last >= 0 else ''}{last:,.0f}",
            ha="right", va="bottom", fontsize=13, fontweight="bold", color=LINE, zorder=4)

    if payback_years is not None and 0 < payback_years < n:
        cx = float(payback_years)
        ax.annotate(
            f"約{payback_years:.1f}年で回収",
            xy=(cx, 0), xytext=(cx + 1.8, y_lo + yr * 0.5),
            fontsize=11, fontweight="bold", color="#333333", ha="center", va="center",
            bbox=dict(boxstyle="round,pad=0.5", fc="white", ec="#C9CDD3", lw=1.2),
            arrowprops=dict(arrowstyle="-|>", color="#8A94A2", lw=1.4,
                            connectionstyle="arc3,rad=-0.2"),
            zorder=5)

    ax.set_xticks(xs)
    ax.set_xticklabels(labels, fontsize=9, color="#333333")
    ax.tick_params(axis="y", labelsize=9, color="#B9C0CA")
    ax.tick_params(axis="x", length=0)
    for sp in ("top", "right"):
        ax.spines[sp].set_visible(False)
    ax.spines["left"].set_color("#B9C0CA")
    ax.spines["bottom"].set_color("#B9C0CA")
    ax.yaxis.grid(True, color=GRID, linewidth=0.8, zorder=0)
    ax.set_axisbelow(True)
    ax.margins(x=0.03)
    ax.set_title(title, fontsize=11.5, fontweight="bold", color=NAVY, loc="left", pad=10)
    return _fig_to_png(fig, out_path, dpi)


# ═════════════════════════════════════════════════════════════════════
# ⑥ 電力使用量配分ドーナツ（分析方法スライド A）
# ═════════════════════════════════════════════════════════════════════
def make_alloc_donut_png(
    labels: Sequence[str],
    values: Sequence[float],
    *,
    colors: Sequence[str] | None = None,
    highlight_index: int | None = None,
    title: str = "電力使用量配分（製造拠点モデル）",
    out_path: str | None = None,
    dpi: int = 200,
    fig_w: float = 5.7,
    fig_h: float = 3.9,
) -> bytes:
    """電力使用量配分のドーナツ。highlight_index の要素を緑＋枠付きラベルで強調。"""
    import math
    set_jp_font()
    vals = [float(v) for v in values]
    labs = list(labels)
    n = len(vals)
    if colors is None:
        base = ["#13315C", "#3E5C86", "#8FA3BF", "#C3CCD9", "#6B7C93"]
        cols = [base[i % len(base)] for i in range(n)]
    else:
        cols = list(colors)
    if highlight_index is not None and 0 <= highlight_index < n:
        cols[highlight_index] = GREEN

    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    wedges, _ = ax.pie(vals, colors=cols, startangle=90, counterclock=False,
                       wedgeprops=dict(width=0.40, edgecolor="white", linewidth=2))
    ax.set_aspect("equal")
    for i, (w, lab, val) in enumerate(zip(wedges, labs, vals)):
        ang = math.radians((w.theta1 + w.theta2) / 2.0)
        x, y = math.cos(ang), math.sin(ang)
        ha = "left" if x >= 0 else "right"
        is_hi = (i == highlight_index)
        t = ax.annotate(f"{lab}：{val:.1f}%",
                        xy=(x * 0.99, y * 0.99), xytext=(x * 1.22, y * 1.30),
                        ha=ha, va="center", fontsize=10,
                        color=(GREEN if is_hi else "#333333"),
                        fontweight=("bold" if is_hi else "normal"),
                        arrowprops=dict(arrowstyle="-", color="#BBBBBB", lw=0.8))
        if is_hi:
            t.set_bbox(dict(boxstyle="round,pad=0.3", fc="white", ec=GREEN, lw=1.4))
    ax.set_title(title, fontsize=11.5, fontweight="bold", color=NAVY, loc="left", pad=6)
    return _fig_to_png(fig, out_path, dpi)


# ═════════════════════════════════════════════════════════════════════
# ⑦ ピーク需要カットの模式図（分析方法スライド C）
# ═════════════════════════════════════════════════════════════════════
def make_peak_curve_png(
    *,
    threshold: float = 0.60,
    out_path: str | None = None,
    dpi: int = 200,
    fig_w: float = 5.3,
    fig_h: float = 1.95,
) -> bytes:
    """デマンド曲線＋ピークカット閾値（緑破線）＋削減量（山の頭）を示す模式図。"""
    import math
    set_jp_font()
    N = 180
    xs = [-3.0 + 6.0 * i / (N - 1) for i in range(N)]
    ys = [math.exp(-(x * x) / 0.85) + 0.12 * math.exp(-((x - 1.6) ** 2) / 0.5) for x in xs]
    ymax = max(ys)
    ys = [y / ymax for y in ys]

    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    ax.plot(xs, ys, color=NAVY, lw=2.2, zorder=3)
    ax.fill_between(xs, [threshold] * N, ys,
                    where=[y > threshold for y in ys],
                    color=GREEN, alpha=0.22, zorder=2)
    ax.axhline(threshold, color=GREEN, ls=(0, (5, 3)), lw=1.6, zorder=3)
    ax.text(xs[0], threshold + 0.04, "ピークカット閾値", color=GREEN,
            fontsize=9, fontweight="bold", va="bottom", ha="left")
    ax.annotate("", xy=(0.32, 1.0), xytext=(0.32, threshold),
                arrowprops=dict(arrowstyle="<->", color="#555555", lw=1.3), zorder=4)
    ax.text(0.5, (1.0 + threshold) / 2, "削減量", color="#333333",
            fontsize=9, fontweight="bold", va="center", ha="left")
    ax.set_ylim(0, 1.18)
    ax.set_xlim(xs[0], xs[-1])
    ax.axis("off")
    return _fig_to_png(fig, out_path, dpi)


# ═════════════════════════════════════════════════════════════════════
# ⑧ ESGアイコン（緑シルエット：CO2削減／樹木／住宅）
# ═════════════════════════════════════════════════════════════════════
def make_esg_icon_png(
    kind: str,
    *,
    color: str = GREEN,
    out_path: str | None = None,
    dpi: int = 200,
    size: float = 1.55,
) -> bytes:
    """kind: 'co2'（雲＋下矢印）/ 'tree'（針葉樹）/ 'house'（家）の緑シルエットアイコン。"""
    from matplotlib.patches import Circle, Polygon as Poly, Rectangle as Rect
    set_jp_font()
    fig, ax = plt.subplots(figsize=(size, size))
    ax.set_xlim(0, 10); ax.set_ylim(0, 10); ax.set_aspect("equal"); ax.axis("off")
    if kind == "co2":
        for cx, cy, r in [(3.5, 6.0, 1.7), (5.3, 6.7, 2.1), (7.1, 6.0, 1.7)]:
            ax.add_patch(Circle((cx, cy), r, fc=color, ec="none"))
        ax.add_patch(Rect((3.1, 4.6), 4.6, 2.2, fc=color, ec="none"))
        ax.annotate("", xy=(5.3, 4.9), xytext=(5.3, 7.0),
                    arrowprops=dict(arrowstyle="-|>", color="white", lw=3.4, mutation_scale=22))
    elif kind == "tree":
        ax.add_patch(Rect((4.55, 1.2), 0.9, 1.7, fc=color, ec="none"))
        for by, hw, ht in [(2.5, 2.7, 2.3), (4.1, 2.2, 2.1), (5.6, 1.7, 2.0)]:
            ax.add_patch(Poly([(5 - hw, by), (5 + hw, by), (5, by + ht)], fc=color, ec="none"))
    elif kind == "house":
        ax.add_patch(Rect((3.1, 1.6), 3.8, 3.3, fc=color, ec="none"))
        ax.add_patch(Poly([(2.4, 4.8), (7.6, 4.8), (5.0, 7.6)], fc=color, ec="none"))
        ax.add_patch(Rect((4.4, 1.6), 1.2, 2.0, fc="white", ec="none"))
    return _fig_to_png(fig, out_path, dpi)


# ═════════════════════════════════════════════════════════════════════
# デモ（参考画像の再現）: python slide_image_export.py
# ═════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    months = ["4月", "5月", "6月", "7月", "8月", "9月",
              "10月", "11月", "12月", "1月", "2月", "3月"]
    usage = [404856, 420342, 457992, 530358, 530598, 518052,
             417498, 363138, 333762, 176316, 271482, 341706]
    reduction = [6367, 6798, 7847, 9863, 9870, 9520,
                 6719, 5205, 4386, 0, 2651, 4607]
    demand = [1386, 1392, 1530, 1632, 1584, 1518,
              1506, 1296, 1218, 1158, 1296, 1230]

    make_demand_chart_png(months, usage, reduction, demand,
                          target_units=17, out_path="_demo_chart.png")
    print("wrote _demo_chart.png")

    items = [
        {"系統名": "女子更衣室エアコン", "設置場所": "工場屋外", "メーカー": "三菱電機", "型式": "PUZ-ZRP80HA11", "制御可否": "○"},
        {"系統名": "男子更衣室エアコン", "設置場所": "工場屋外", "メーカー": "三菱電機", "型式": "PUZ-ZRMP56HA15", "制御可否": "×"},
        {"系統名": "社長室", "設置場所": "工場屋外", "メーカー": "三菱電機", "型式": "PUZ-ZRP80HA10", "制御可否": "○"},
        {"系統名": "社長室", "設置場所": "工場屋外", "メーカー": "三菱電機", "型式": "PUZ-ZRMP80HA12", "制御可否": "○"},
        {"系統名": "1F休憩室", "設置場所": "工場屋外", "メーカー": "三菱電機", "型式": "PUSY-FFP280MH1", "制御可否": "×"},
        {"系統名": "事務エアコン", "設置場所": "工場屋外", "メーカー": "三菱電機", "型式": "PUZ-ZRMP160KA", "制御可否": "○"},
        {"系統名": "食堂西側(品管側)", "設置場所": "工場屋外", "メーカー": "三菱電機", "型式": "PUZ-ZRMP63KA", "制御可否": "○"},
        {"系統名": "応接室エアコン", "設置場所": "工場屋外", "メーカー": "三菱電機", "型式": "PUZ-ERP224KA3", "制御可否": "○"},
        {"系統名": "品質管理室", "設置場所": "工場屋外", "メーカー": "三菱電機", "型式": "PUZ-ZRP140KA5", "制御可否": "○"},
        {"系統名": "会議室エアコン", "設置場所": "工場屋外", "メーカー": "三菱電機", "型式": "PUZ-ZRMP160KA", "制御可否": "○"},
        {"系統名": "食堂束側(道路側)", "設置場所": "工場屋外", "メーカー": "三菱電機", "型式": "PUZ-ZRMP112KA", "制御可否": "○"},
        {"系統名": "PR室", "設置場所": "工場屋外", "メーカー": "三菱電機", "型式": "MPUZ-P80HA3", "制御可否": "○"},
        {"系統名": "不明", "設置場所": "工場屋外", "メーカー": "三菱電機", "型式": "PUZ-ZRMP112KA2", "制御可否": "○"},
        {"系統名": "工作室", "設置場所": "工場屋外", "メーカー": "三菱電機", "型式": "PUZ-ZRMP80HA12", "制御可否": "○"},
        {"系統名": "応接室", "設置場所": "工場屋外", "メーカー": "三菱電機", "型式": "PUZ-ERP50KA7", "制御可否": "○"},
        {"系統名": "生産管理室", "設置場所": "工場屋外", "メーカー": "三菱電機", "型式": "PUZ-ERP56KA7", "制御可否": "○"},
        {"系統名": "不明", "設置場所": "工場屋外", "メーカー": "三菱電機", "型式": "MPUZ-ERP50KA2", "制御可否": "○"},
        {"系統名": "包餡室西側", "設置場所": "工場屋外", "メーカー": "三菱電機", "型式": "PUHV-P280DM-E", "制御可否": "×"},
        {"系統名": "包餡室東側", "設置場所": "工場屋外", "メーカー": "三菱電機", "型式": "PUHV-P280DM-E", "制御可否": "×"},
        {"系統名": "成形機室(南側)", "設置場所": "工場屋外", "メーカー": "三菱電機", "型式": "MPUZ-ERP2800KA", "制御可否": "○"},
        {"系統名": "チラーNo.1", "設置場所": "空調用冷凍機室", "メーカー": "ダイキン", "型式": "RCF2000WVTC", "制御可否": "×"},
        {"系統名": "チラーNo.2", "設置場所": "空調用冷凍機室", "メーカー": "ダイキン", "型式": "RCF2000WVTC", "制御可否": "×"},
        {"系統名": "チラーNo.3", "設置場所": "空調用冷凍機室", "メーカー": "ダイキン", "型式": "RCF2000WVTC", "制御可否": "×"},
        {"系統名": "チラーNo.4", "設置場所": "空調用冷凍機室", "メーカー": "ダイキン", "型式": "RCF2000WVTC", "制御可否": "×"},
        {"系統名": "チラーNo.5", "設置場所": "空調用冷凍機室", "メーカー": "ダイキン", "型式": "RCF2000WVTC", "制御可否": "×"},
        {"系統名": "チラーNo.6", "設置場所": "空調用冷凍機室", "メーカー": "ダイキン", "型式": "RCF2000WVTC", "制御可否": "×"},
    ]
    # 27件 → 1枚に収まる
    pages = make_control_list_pngs(items, total_units=27, controllable_units=17,
                                   out_path_prefix="_demo_list")
    print(f"wrote {len(pages)} page(s): _demo_list_1.png ...")

    # 分割デモ：120件 → 自動で複数ページ（50件/枚）
    big = []
    for i in range(120):
        src = items[i % len(items)]
        big.append({**src, "系統名": f"{src['系統名']}{i+1:03d}"})
    big_pages = make_control_list_pngs(big, out_path_prefix="_demo_list_big")
    print(f"wrote {len(big_pages)} page(s): _demo_list_big_1.png ...")

    # ③ サマリKPIカード
    invest, gross, net, sys_fee = 3510000, 1611528, 1441728, 169800
    payback = invest / net if net > 0 else 0
    cards = [
        {"icon": "money.png", "title": "初期導入費用", "value": f"{invest:,}円",
         "subtitle": "税抜 導入費用 / 機器代・工事費含む", "color": NAVY},
        {"icon": "graf.png", "title": "年間総削減額（グロス）", "value": f"{gross:,}円",
         "subtitle": "基本料金＋電力量の年間削減合計", "color": GREEN},
        {"icon": "plus.png", "title": "年間実質利点（手残り）", "value": f"{net:,}円",
         "subtitle": f"削減合計 {gross:,}円 − 年間維持費 {sys_fee:,}円", "color": GREEN},
        {"icon": "clock.png", "title": "投資回収期間（ROI）", "value": f"約{payback:.1f}年",
         "subtitle": f"約{round(payback*12)}ヶ月で完全回収、以降は純利益", "color": GREEN},
    ]
    make_summary_cards_png(cards, out_path="_demo_cards.png")
    print("wrote _demo_cards.png")
