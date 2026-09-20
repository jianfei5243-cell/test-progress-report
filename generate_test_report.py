#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
OS4.0 第一阶段全功能测试进度自动化汇总脚本

功能：
  1. 自动发现最新的「Test导出_*.xlsx」文件（默认在 ~/Downloads 与脚本目录查找）
  2. 汇总总体进度、按功能模块(Folder)、按测试批次(Run Name)、按优先级、Top缺陷
  3. 输出：
     - 测试报告_<日期>.html         自包含网页版测试报告（可直接浏览器打开/打印为PDF）
     - 测试进度汇总_<日期>.xlsx     数据汇总表（含原生图表）
     - 测试进度_模块执行率_<日期>.png  模块执行率柱状图

依赖：openpyxl、Pillow（工作区已内置，普通环境可用 pip install openpyxl pillow）

用法：
  python3 generate_test_report.py                      # 自动发现最新导出文件
  python3 generate_test_report.py --input /path/xxx.xlsx
  python3 generate_test_report.py --out-dir ./report
"""
import argparse
import base64
import datetime
import glob
import os
import re
import sys
from collections import Counter, defaultdict

import openpyxl
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

try:
    from PIL import Image, ImageDraw, ImageFont
    HAS_PIL = True
except Exception:
    HAS_PIL = False

# 源文件列索引（0-based），与导出模板固定字段对应
COL = dict(
    folder=2, priority=11, run=33, status=35,
    est_exec=3, est_mh=4, exec_time=37, mh_time=38, defects=42,
)

STATUS_ORDER = ["Passed", "Failed", "N/A", "Retest", "Untested"]
STATUS_LABEL = {
    "Passed": "通过", "Failed": "失败", "N/A": "不适用",
    "Retest": "待复测", "Untested": "未执行",
}

FONT_CJK = "/System/Library/Fonts/Supplemental/Arial Unicode.ttf"


# --------------------------------------------------------------------------- #
# 输入定位与解析
# --------------------------------------------------------------------------- #
def find_latest_input(extra_dirs):
    dirs = [os.path.expanduser("~/Downloads"), os.path.dirname(os.path.abspath(__file__))]
    dirs += list(extra_dirs)
    files = []
    for d in dirs:
        if not d:
            continue
        files += glob.glob(os.path.join(d, "Test导出_*.xlsx"))
    if not files:
        return None
    return max(files, key=os.path.getmtime)


def parse_minutes(v):
    if v is None:
        return 0.0
    s = str(v).strip()
    if not s or s == "None":
        return 0.0
    m = re.match(r"^([\d.]+)\s*(?:m|min)?$", s, re.I)
    if m:
        return float(m.group(1))
    m = re.match(r"^([\d.]+)\s*h$", s, re.I)
    if m:
        return float(m.group(1)) * 60
    return 0.0


def load_and_aggregate(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]

    mod = defaultdict(Counter)     # 功能模块(文件夹一级) -> 状态计数
    run = defaultdict(Counter)     # 测试批次 -> 状态计数
    pri = defaultdict(Counter)     # 优先级 -> 状态计数
    defects = Counter()
    overall = Counter()

    total = 0
    est_exec = est_mh = exec_t = mh_t = 0.0

    for r in ws.iter_rows(min_row=2, values_only=True):
        total += 1
        st = str(r[COL["status"]])
        overall[st] += 1

        f = str(r[COL["folder"]])
        rn = str(r[COL["run"]])
        p = str(r[COL["priority"]])
        if f and f != "None":
            mod[f.split(" > ")[0]][st] += 1
        if rn and rn != "None":
            run[rn][st] += 1
        if p and p != "None":
            pri[p][st] += 1

        est_exec += parse_minutes(r[COL["est_exec"]])
        est_mh += parse_minutes(r[COL["est_mh"]])
        exec_t += parse_minutes(r[COL["exec_time"]])
        mh_t += parse_minutes(r[COL["mh_time"]])

        if st == "Failed":
            d = str(r[COL["defects"]])
            if d and d != "None":
                for did in re.split(r"[;,，\s]+", d):
                    did = did.strip()
                    if did:
                        defects[did] += 1

    wb.close()

    return {
        "total": total,
        "overall": overall,
        "mod": mod,
        "run": run,
        "pri": pri,
        "defects": defects,
        "minutes": {"est_exec": est_exec, "est_mh": est_mh,
                    "exec": exec_t, "mh": mh_t},
    }


def agg(c):
    tot = sum(c.values())
    unt = c.get("Untested", 0)
    ex = tot - unt
    ps = c.get("Passed", 0)
    fl = c.get("Failed", 0)
    na = c.get("N/A", 0)
    rt = c.get("Retest", 0)
    return {
        "总": tot, "未执行": unt, "已执行": ex,
        "通过": ps, "失败": fl, "N/A": na, "Retest": rt,
        "执行率": round(ex / tot * 100, 1) if tot else 0.0,
        "通过率": round(ps / ex * 100, 1) if ex else None,
    }


def sorted_agg(mapping):
    return sorted(mapping.items(), key=lambda x: -sum(x[1].values()))


# --------------------------------------------------------------------------- #
# 输出：PNG 柱状图
# --------------------------------------------------------------------------- #
def render_png(items, title="功能模块执行率（Top15）"):
    if not HAS_PIL:
        return None
    top = 15
    items = items[:top]
    W, row_h, top_m, bot_m = 1400, 54, 92, 50
    H = top_m + len(items) * row_h + bot_m
    img = Image.new("RGB", (W, H), (255, 255, 255))
    dr = ImageDraw.Draw(img)

    font = FONT_CJK if os.path.exists(FONT_CJK) else None
    def loadf(s):
        return ImageFont.truetype(font, s) if font else ImageFont.load_default()

    title_f, lab_f, num_f = loadf(30), loadf(21), loadf(20)
    dr.text((40, 30), title, font=title_f, fill="#1F2937")

    label_x, label_w = 40, 300
    bar_x0 = label_x + label_w + 20
    bar_w = W - bar_x0 - 360
    bar_h = 22

    def trunc(t, f, maxw):
        if dr.textlength(t, font=f) <= maxw:
            return t
        while t and dr.textlength(t + "…", font=f) > maxw:
            t = t[:-1]
        return t + "…"

    for i, (name, d) in enumerate(items):
        cy = top_m + i * row_h + row_h // 2
        lab = trunc(name, lab_f, label_w)
        tw = dr.textlength(lab, font=lab_f)
        dr.text((label_x + label_w - tw, cy - lab_f.size / 2), lab, font=lab_f, fill="#111827")

        rate = d["执行率"]
        x1 = bar_x0 + int(rate / 100 * bar_w)
        color = "#16A34A" if rate >= 50 else ("#F59E0B" if rate >= 20 else "#9CA3AF")
        dr.rectangle([bar_x0, cy - bar_h // 2, x1, cy + bar_h // 2], fill=color)
        dr.text((bar_x0 + bar_w + 18, cy - num_f.size / 2),
                f'{d["已执行"]}/{d["总"]}   {rate}%', font=num_f, fill="#374151")

    import io
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


# --------------------------------------------------------------------------- #
# 输出：Excel 汇总
# --------------------------------------------------------------------------- #
def build_excel(path, data):
    overall = data["overall"]
    total = data["total"]
    unt = overall.get("Untested", 0)
    ex = total - unt
    ps = overall.get("Passed", 0)
    fl = overall.get("Failed", 0)
    na = overall.get("N/A", 0)
    rt = overall.get("Retest", 0)

    wb = Workbook()
    hdr_fill = PatternFill("solid", fgColor="2F5597")
    hdr_font = Font(bold=True, color="FFFFFF")
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws1 = wb.active
    ws1.title = "总体进度"
    summary = [
        ["指标", "数值"],
        ["用例总数", total],
        ["未执行 Untested", unt],
        ["已执行", ex],
        ["执行覆盖率", f"{ex/total:.1%}"],
        ["通过 Passed", ps],
        ["失败 Failed", fl],
        ["不适用 N/A", na],
        ["待复测 Retest", rt],
        ["通过率（已执行口径）", f"{ps/ex:.1%}"],
        ["通过率（剔除N/A口径）", f"{ps/(ps+fl+rt):.1%}"],
        ["失败率（已执行口径）", f"{fl/ex:.1%}"],
    ]
    for i, row in enumerate(summary, 1):
        for j, v in enumerate(row, 1):
            c = ws1.cell(i, j, v)
            c.border = border
            if i == 1:
                c.fill = hdr_fill
                c.font = hdr_font
                c.alignment = Alignment(horizontal="center")
    ws1.column_dimensions["A"].width = 28
    ws1.column_dimensions["B"].width = 16

    def write_table(ws, mapping):
        headers = ["模块", "用例总数", "未执行", "通过", "失败", "N/A", "Retest", "执行率%", "通过率%"]
        for j, h in enumerate(headers, 1):
            c = ws.cell(1, j, h)
            c.fill = hdr_fill
            c.font = hdr_font
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = border
        rows = sorted_agg(mapping)
        for i, (name, cnt) in enumerate(rows, 2):
            d = agg(cnt)
            vals = [name, d["总"], d["未执行"], d["通过"], d["失败"], d["N/A"], d["Retest"],
                    d["执行率"], d["通过率"]]
            for j, v in enumerate(vals, 1):
                c = ws.cell(i, j, v)
                c.border = border
                if j >= 2:
                    c.alignment = Alignment(horizontal="center")
                if j == 9 and v is None:
                    c.value = "-"
        for j, w in enumerate([30, 12, 12, 12, 12, 10, 10, 12, 12], 1):
            ws.column_dimensions[get_column_letter(j)].width = w
        ws.freeze_panes = "A2"
        return len(rows) + 1

    ws2 = wb.create_sheet("功能模块")
    n2 = write_table(ws2, data["mod"])
    ws3 = wb.create_sheet("测试批次")
    write_table(ws3, data["run"])

    topn = 15
    ch1 = BarChart()
    ch1.type = "bar"
    ch1.style = 10
    ch1.title = "功能模块执行率（Top15）"
    ch1.y_axis.title = "模块"
    ch1.x_axis.title = "执行率 %"
    ch1.add_data(Reference(ws2, min_col=8, min_row=1, max_row=min(topn + 1, n2)), titles_from_data=True)
    ch1.set_categories(Reference(ws2, min_col=1, min_row=2, max_row=min(topn + 1, n2)))
    ch1.width, ch1.height = 20, 11
    ws2.add_chart(ch1, "K2")

    ch2 = BarChart()
    ch2.type = "bar"
    ch2.grouping = "stacked"
    ch2.overlap = 100
    ch2.title = "功能模块状态分布（Top15）"
    ch2.y_axis.title = "模块"
    ch2.x_axis.title = "用例数"
    ch2.add_data(Reference(ws2, min_col=3, max_col=7, min_row=1, max_row=min(topn + 1, n2)), titles_from_data=True)
    ch2.set_categories(Reference(ws2, min_col=1, min_row=2, max_row=min(topn + 1, n2)))
    ch2.width, ch2.height = 20, 11
    ws2.add_chart(ch2, "K22")

    wb.save(path)


# --------------------------------------------------------------------------- #
# 输出：HTML 测试报告
# --------------------------------------------------------------------------- #
def esc(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def pct(n):
    return f"{n:.1f}%" if n is not None else "-"


def build_html(path, data, png_bytes, source_name, source_date):
    overall = data["overall"]
    total = data["total"]
    unt = overall.get("Untested", 0)
    ex = total - unt
    ps = overall.get("Passed", 0)
    fl = overall.get("Failed", 0)
    na = overall.get("N/A", 0)
    rt = overall.get("Retest", 0)
    minutes = data["minutes"]

    cards = [
        ("用例总数", f"{total}"),
        ("已执行", f"{ex}"),
        ("执行覆盖率", f"{ex/total:.1%}"),
        ("通过 Passed", f"{ps}"),
        ("失败 Failed", f"{fl}"),
        ("不适用 N/A", f"{na}"),
        ("待复测 Retest", f"{rt}"),
        ("通过率(已执行)", f"{ps/ex:.1%}"),
        ("通过率(剔除N/A)", f"{ps/(ps+fl+rt):.1%}"),
        ("失败率", f"{fl/ex:.1%}"),
    ]
    cards_html = "".join(
        f'<div class="card"><div class="k">{k}</div><div class="v">{v}</div></div>'
        for k, v in cards
    )

    def table_html(mapping, header):
        rows = sorted_agg(mapping)
        trs = []
        for name, cnt in rows:
            d = agg(cnt)
            trs.append(
                "<tr><td class='l'>{}</td><td>{}</td><td>{}</td><td>{}</td><td>{}</td>"
                "<td>{}</td><td>{}</td><td>{}</td><td>{}</td></tr>".format(
                    esc(name), d["总"], d["未执行"], d["通过"], d["失败"],
                    d["N/A"], d["Retest"], pct(d["执行率"]), pct(d["通过率"]),
                )
            )
        return f"""
        <table>
          <thead><tr>{''.join(f'<th>{h}</th>' for h in header)}</tr></thead>
          <tbody>{''.join(trs)}</tbody>
        </table>"""

    pri_header = ["优先级", "用例总数", "未执行", "通过", "失败", "N/A", "Retest", "执行率%", "通过率%"]
    pri_html = table_html(data["pri"], pri_header)
    mod_html = table_html(data["mod"], ["模块", "用例总数", "未执行", "通过", "失败", "N/A", "Retest", "执行率%", "通过率%"])
    run_html = table_html(data["run"], ["测试批次", "用例总数", "未执行", "通过", "失败", "N/A", "Retest", "执行率%", "通过率%"])

    defects = data["defects"].most_common(15)
    defects_html = "".join(
        f"<li><code>{esc(d)}</code><span>{n} 条</span></li>" for d, n in defects
    )

    chart_html = ""
    if png_bytes:
        b64 = base64.b64encode(png_bytes).decode("ascii")
        chart_html = f'<img class="chart" src="data:image/png;base64,{b64}" alt="模块执行率"/>'

    html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>OS4.0 第一阶段全功能测试报告</title>
<style>
:root {{ --blue:#2F5597; --line:#e5e7eb; --bg:#f6f7f9; }}
* {{ box-sizing:border-box; }}
body {{ margin:0; font-family:-apple-system,"PingFang SC","Microsoft YaHei",Arial,sans-serif; color:#1f2937; background:var(--bg); }}
.wrap {{ max-width:1180px; margin:0 auto; padding:32px 24px 64px; }}
h1 {{ font-size:26px; margin:0 0 4px; }}
.sub {{ color:#6b7280; font-size:13px; margin-bottom:24px; }}
h2 {{ font-size:18px; margin:36px 0 12px; padding-left:10px; border-left:4px solid var(--blue); }}
.cards {{ display:grid; grid-template-columns:repeat(5,1fr); gap:12px; }}
.card {{ background:#fff; border:1px solid var(--line); border-radius:10px; padding:14px 12px; }}
.card .k {{ font-size:12px; color:#6b7280; }}
.card .v {{ font-size:22px; font-weight:700; margin-top:6px; color:var(--blue); }}
table {{ width:100%; border-collapse:collapse; background:#fff; font-size:13px; }}
th,td {{ border:1px solid var(--line); padding:7px 8px; text-align:center; }}
th {{ background:var(--blue); color:#fff; font-weight:600; }}
td.l {{ text-align:left; font-weight:600; white-space:nowrap; }}
tbody tr:nth-child(even) {{ background:#f8fafc; }}
.chart {{ max-width:100%; margin:12px 0; border:1px solid var(--line); border-radius:10px; background:#fff; }}
ul.def {{ list-style:none; padding:0; margin:0; }}
ul.def li {{ background:#fff; border:1px solid var(--line); border-radius:8px; padding:8px 12px; margin-bottom:6px; display:flex; justify-content:space-between; font-size:13px; }}
ul.def code {{ color:#b91c1c; }}
.info {{ background:#fff; border-left:4px solid var(--blue); border-radius:8px; padding:12px 16px; font-size:13px; line-height:1.8; }}
.risk {{ background:#fff; border-left:4px solid #b91c1c; border-radius:8px; padding:12px 16px; font-size:13px; line-height:1.8; }}
@media print {{ body{{background:#fff}} .card,.chart,table,ul.def li,.risk{{box-shadow:none}} }}
</style>
</head>
<body>
<div class="wrap">
  <h1>【P62】OS4.0 第一阶段全功能测试报告</h1>
  <div class="sub">数据源：{esc(source_name)} ｜ 生成时间：{esc(source_date)} ｜ 用例总数：{total}</div>

  <div class="cards">{cards_html}</div>

  <h2>模块执行率（Top15）</h2>
  {chart_html}

  <h2>功能模块进度（Folder 一级）</h2>
  {mod_html}

  <h2>测试批次进度（Run Name）</h2>
  {run_html}

  <h2>优先级进度</h2>
  {pri_html}

  <h2>Top 缺陷</h2>
  <ul class="def">{defects_html}</ul>

  <h2>人力投入</h2>
  <div class="info">
    估算执行时长：{minutes['est_exec']:.0f} 分钟 ｜ 实际执行时长：{minutes['exec']:.0f} 分钟<br/>
    估算人时：{minutes['est_mh']:.0f} 分钟 ｜ 实际人时：{minutes['mh']:.0f} 分钟
  </div>
</div>
</body>
</html>"""

    with open(path, "w", encoding="utf-8") as fh:
        fh.write(html)


# --------------------------------------------------------------------------- #
# 主流程
# --------------------------------------------------------------------------- #
def main(argv=None):
    ap = argparse.ArgumentParser(description="OS4.0 测试进度自动化汇总")
    ap.add_argument("--input", "-i", help="导出 xlsx 路径（缺省自动查找最新文件）")
    ap.add_argument("--out-dir", "-o", default=os.path.dirname(os.path.abspath(__file__)),
                    help="输出目录（缺省为脚本所在目录）")
    ap.add_argument("--search", "-s", nargs="*", default=[], help="额外搜索目录")
    args = ap.parse_args(argv)

    src = args.input or find_latest_input(args.search)
    if not src or not os.path.exists(src):
        print("未找到 Test导出_*.xlsx 文件，请用 --input 指定。", file=sys.stderr)
        return 2

    os.makedirs(args.out_dir, exist_ok=True)
    stamp = datetime.datetime.fromtimestamp(os.path.getmtime(src)).strftime("%Y%m%d")
    prefix = os.path.join(args.out_dir, f"{stamp}")

    print("解析源文件：", src)
    data = load_and_aggregate(src)

    items = [(name, agg(cnt)) for name, cnt in sorted_agg(data["mod"])]
    png_bytes = render_png(items)

    xlsx_path = f"{prefix}_测试进度汇总.xlsx"
    png_path = f"{prefix}_测试进度_模块执行率.png"
    html_path = f"{prefix}_测试报告.html"

    build_excel(xlsx_path, data)

    if png_bytes:
        with open(png_path, "wb") as fh:
            fh.write(png_bytes)

    build_html(html_path, data, png_bytes, os.path.basename(src),
               datetime.datetime.fromtimestamp(os.path.getmtime(src)).strftime("%Y-%m-%d %H:%M"))

    o = data["overall"]
    total = data["total"]
    ex = total - o.get("Untested", 0)
    print("=" * 46)
    print(f"用例总数：{total}  已执行：{ex}（{ex/total:.1%}）")
    print(f"通过：{o.get('Passed',0)}  失败：{o.get('Failed',0)}  "
          f"N/A：{o.get('N/A',0)}  Retest：{o.get('Retest',0)}")
    print("输出文件：")
    for p in (html_path, xlsx_path, png_path):
        print("  -", p)
    return 0


if __name__ == "__main__":
    sys.exit(main())
