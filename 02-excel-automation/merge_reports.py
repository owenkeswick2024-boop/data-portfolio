"""
多门店、多月份 Excel 报表 → 一键清洗合并 → 带公式和格式的汇总表

用法：
    python merge_reports.py                      # 默认读取 input/，输出 output/销售汇总.xlsx
    python merge_reports.py -i 你的文件夹 -o 结果.xlsx

文件名约定：<门店名>_<年>年<月>月.xlsx，例如 “深圳福田店_2026年7月.xlsx”
"""
import argparse
import re
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ---------- 1. 列名映射：把各门店五花八门的列名统一成标准名 ----------
COLUMN_ALIASES = {
    "单号": ["单号", "订单号", "流水号", "order id"],
    "日期": ["日期", "销售日期", "交易日期", "date"],
    "商品": ["商品", "商品名称", "品名", "产品", "product"],
    "数量": ["数量", "销量", "件数", "qty"],
    "单价": ["单价", "单价(元)", "单价（元）", "price"],
    "金额": ["金额", "销售额", "金额(元)", "金额（元）", "amount"],
}
ALIAS_TO_STD = {a.lower(): std for std, aliases in COLUMN_ALIASES.items() for a in aliases}
FILENAME_RE = re.compile(r"(?P<store>.+?)_(?P<year>\d{4})年(?P<month>\d{1,2})月")


def find_header_row(raw: pd.DataFrame) -> int:
    """在前 10 行里找到真正的表头（跳过标题行、空行）。"""
    for i in range(min(10, len(raw))):
        cells = {str(v).strip().lower() for v in raw.iloc[i] if pd.notna(v)}
        if len(cells & set(ALIAS_TO_STD)) >= 3:
            return i
    raise ValueError("前 10 行内没有找到表头")


def to_number(s: pd.Series) -> pd.Series:
    """'1,280.00' 这类文本金额 → 数字。"""
    return pd.to_numeric(s.astype(str).str.replace(",", "", regex=False).str.strip(), errors="coerce")


def clean_one_file(path: Path):
    m = FILENAME_RE.search(path.stem)
    if not m:
        raise ValueError(f"文件名不符合约定：{path.name}")
    store, month = m["store"], f"{m['year']}-{int(m['month']):02d}"

    raw = pd.read_excel(path, header=None)
    h = find_header_row(raw)
    df = raw.iloc[h + 1:].copy()
    df.columns = [ALIAS_TO_STD.get(str(c).strip().lower(), str(c).strip()) for c in raw.iloc[h]]
    missing = set(COLUMN_ALIASES) - set(df.columns)
    if missing:
        raise ValueError(f"{path.name} 缺少列：{missing}")
    df = df[list(COLUMN_ALIASES)]

    log = {"文件": path.name, "门店": store, "月份": month, "原始行数": len(df)}

    # 空行
    before = len(df)
    df = df.dropna(how="all")
    log["删除空行"] = before - len(df)

    # 手工合计行
    is_total = df["单号"].astype(str).str.contains("合计|小计|总计")
    log["删除合计行"] = int(is_total.sum())
    df = df[~is_total]

    # 统一格式
    df["单号"] = df["单号"].astype(str).str.strip()
    df["商品"] = df["商品"].astype(str).str.strip()
    df["日期"] = pd.to_datetime(df["日期"].astype(str).str.split(" ").str[0], format="mixed", errors="coerce")
    for col in ["数量", "单价", "金额"]:
        df[col] = to_number(df[col])

    bad = df[["日期", "数量", "单价", "金额"]].isna().any(axis=1)
    log["无法识别的行"] = int(bad.sum())
    df = df[~bad]

    # 金额校验：金额应等于 数量 × 单价
    mismatch = (df["数量"] * df["单价"] - df["金额"]).abs() > 0.01
    log["金额不符(已按数量×单价修正)"] = int(mismatch.sum())
    df.loc[mismatch, "金额"] = df.loc[mismatch, "数量"] * df.loc[mismatch, "单价"]

    # 重复录入：同一单号出现多次只保留一条（注意不能按整行去重——
    # 同一天卖出同款同数量的两笔订单是真实存在的）
    before = len(df)
    df = df.drop_duplicates(subset="单号", keep="first")
    log["删除重复行"] = before - len(df)

    log["有效行数"] = len(df)
    df.insert(0, "月份", month)
    df.insert(0, "门店", store)
    return df, log


# ---------- 2. 输出样式 ----------
FONT = "Microsoft YaHei"
HEAD_FILL = PatternFill("solid", fgColor="1F4E78")
TOTAL_FILL = PatternFill("solid", fgColor="DDEBF7")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
MONEY = '#,##0;-#,##0;"-"'


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row, c)
        cell.font = Font(name=FONT, bold=True, color="FFFFFF")
        cell.fill = HEAD_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER


def style_body(ws, r1, r2, ncols):
    for r in range(r1, r2 + 1):
        for c in range(1, ncols + 1):
            cell = ws.cell(r, c)
            cell.font = Font(name=FONT, bold=cell.font.bold)
            cell.border = BORDER


def set_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def write_workbook(data: pd.DataFrame, logs: pd.DataFrame, out: Path):
    wb = Workbook()
    stores = list(dict.fromkeys(data["门店"]))
    months = sorted(data["月份"].unique())
    n = len(data)

    # --- 明细 ---
    det = wb.active
    det.title = "明细"
    cols = ["门店", "月份", "单号", "日期", "商品", "数量", "单价", "金额"]
    det.append(cols)
    for row in data[cols].itertuples(index=False):
        det.append(list(row))
    style_header(det, 1, len(cols))
    style_body(det, 2, n + 1, len(cols))
    for r in range(2, n + 2):
        det.cell(r, 4).number_format = "yyyy-mm-dd"
        det.cell(r, 7).number_format = MONEY
        det.cell(r, 8).number_format = MONEY
    det.freeze_panes = "A2"
    det.auto_filter.ref = f"A1:H{n + 1}"
    set_widths(det, [14, 10, 12, 12, 16, 8, 10, 12])
    rng = lambda col: f"明细!${col}$2:${col}${n + 1}"

    # --- 汇总看板（全部用公式，改明细会自动更新）---
    sm = wb.create_sheet("门店月度汇总", 0)
    sm["A1"] = "各门店月度销售额（元）"
    sm["A1"].font = Font(name=FONT, bold=True, size=14)
    hdr = 3
    sm.cell(hdr, 1, "门店")
    for j, mth in enumerate(months, 2):
        sm.cell(hdr, j, mth)
    tot_col = len(months) + 2
    sm.cell(hdr, tot_col, "合计")
    sm.cell(hdr, tot_col + 1, "占比")
    style_header(sm, hdr, tot_col + 1)

    first, last = hdr + 1, hdr + len(stores)
    for i, s in enumerate(stores, first):
        sm.cell(i, 1, s)
        for j in range(2, tot_col):
            ml = get_column_letter(j)
            sm.cell(i, j, f'=SUMIFS({rng("H")},{rng("A")},$A{i},{rng("B")},{ml}${hdr})').number_format = MONEY
        L = get_column_letter(tot_col - 1)
        sm.cell(i, tot_col, f"=SUM(B{i}:{L}{i})").number_format = MONEY
        T = get_column_letter(tot_col)
        sm.cell(i, tot_col + 1, f"=IF(${T}${last + 1}=0,0,{T}{i}/${T}${last + 1})").number_format = "0.0%"

    tr = last + 1
    sm.cell(tr, 1, "合计")
    for j in range(2, tot_col + 2):
        L = get_column_letter(j)
        sm.cell(tr, j, f"=SUM({L}{first}:{L}{last})").number_format = "0.0%" if j == tot_col + 1 else MONEY
    style_body(sm, first, tr, tot_col + 1)
    for c in range(1, tot_col + 2):
        sm.cell(tr, c).fill = TOTAL_FILL
        sm.cell(tr, c).font = Font(name=FONT, bold=True)

    gr = tr + 1
    sm.cell(gr, 1, "环比增长")
    for j in range(3, tot_col):
        p, c_ = get_column_letter(j - 1), get_column_letter(j)
        sm.cell(gr, j, f"=IF({p}{tr}=0,0,{c_}{tr}/{p}{tr}-1)").number_format = "+0.0%;-0.0%;0.0%"
    style_body(sm, gr, gr, tot_col + 1)
    sm.cell(gr, 1).font = Font(name=FONT, italic=True, color="595959")
    set_widths(sm, [16] + [14] * len(months) + [14, 10])

    chart = BarChart()
    chart.type = "col"
    chart.title = "各门店月度销售额"
    chart.y_axis.title = "元"
    chart.add_data(Reference(sm, min_col=2, max_col=tot_col - 1, min_row=hdr, max_row=last), titles_from_data=True)
    chart.set_categories(Reference(sm, min_col=1, min_row=first, max_row=last))
    chart.legend.position = "b"
    chart.height, chart.width = 8, 15.5
    sm.add_chart(chart, f"A{gr + 3}")

    # --- 商品排行 ---
    pr = wb.create_sheet("商品排行", 1)
    pr.append(["商品", "销量", "销售额", "占比"])
    products = data.groupby("商品")["金额"].sum().sort_values(ascending=False).index
    last_p = len(products) + 1
    for i, p in enumerate(products, 2):
        pr.cell(i, 1, p)
        pr.cell(i, 2, f'=SUMIFS({rng("F")},{rng("E")},$A{i})')
        pr.cell(i, 3, f'=SUMIFS({rng("H")},{rng("E")},$A{i})').number_format = MONEY
        pr.cell(i, 4, f"=IF(SUM($C$2:$C${last_p})=0,0,C{i}/SUM($C$2:$C${last_p}))").number_format = "0.0%"
    style_header(pr, 1, 4)
    style_body(pr, 2, last_p, 4)
    pr.freeze_panes = "A2"
    set_widths(pr, [18, 10, 14, 10])

    # --- 清洗日志 ---
    lg = wb.create_sheet("清洗日志")
    lg.append(list(logs.columns))
    for row in logs.itertuples(index=False):
        lg.append(list(row))
    style_header(lg, 1, len(logs.columns))
    style_body(lg, 2, len(logs) + 1, len(logs.columns))
    set_widths(lg, [26, 14, 10] + [12] * (len(logs.columns) - 3))
    for c in range(1, len(logs.columns) + 1):
        lg.cell(1, c).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    lg.row_dimensions[1].height = 36

    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)


def main():
    here = Path(__file__).parent
    ap = argparse.ArgumentParser(description="合并清洗多门店 Excel 报表")
    ap.add_argument("-i", "--input", default=here / "input", type=Path)
    ap.add_argument("-o", "--output", default=here / "output" / "销售汇总.xlsx", type=Path)
    args = ap.parse_args()

    files = sorted(p for p in args.input.glob("*.xlsx") if not p.name.startswith("~$"))
    if not files:
        raise SystemExit(f"{args.input} 里没有 .xlsx 文件")

    frames, logs = [], []
    for f in files:
        df, log = clean_one_file(f)
        frames.append(df)
        logs.append(log)
        print(f"✓ {f.name}: {log['原始行数']} 行 → {log['有效行数']} 行有效")

    data = pd.concat(frames, ignore_index=True).sort_values(["门店", "日期", "单号"]).reset_index(drop=True)
    logs = pd.DataFrame(logs)
    write_workbook(data, logs, args.output)
    print(f"\n完成：{len(files)} 个文件，{len(data)} 条有效记录 → {args.output}")


if __name__ == "__main__":
    main()
