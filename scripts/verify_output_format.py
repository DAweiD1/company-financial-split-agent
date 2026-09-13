#!/usr/bin/env python3
"""
验证成品 Excel 格式合规性（作业说明书 AGENTS.md 的格式与交叉校验项）。
用法：python verify_output_format.py <path_to_outputs.xlsx>

检查项：
1. 日期行：每年4季度（3/30, 6/30, 9/30, 12/31），格式 YYYY-MM-DD，Arial 11 加粗
2. 数字格式：yoy/毛利率/占比→0.00%，金额→0.00，销量→#,##0.00
3. 空值：Sheet 1 禁止 "—"/"--" 占位符
4. 交叉验证：产品合计≈总收入，地区合计≈总收入
"""
import openpyxl
import sys
from collections import defaultdict

def verify(path):
    wb = openpyxl.load_workbook(path)
    ws = wb["预测模型"]
    errors = []
    warnings = []

    # ---- 1. 日期行格式 ----
    dates = []
    for c in range(2, ws.max_column + 1):
        d = ws.cell(1, c).value
        if d:
            dates.append((c, d))
    years = defaultdict(list)
    for c, d in dates:
        years[d.year].append((c, d.month, d.day))
    for y, cols in sorted(years.items()):
        if len(cols) == 4:
            months = sorted([(m, dd) for _, m, dd in cols])
            expected = [(3, 30), (6, 30), (9, 30), (12, 31)]
            if months != expected:
                errors.append(f"{y}年季度日期不正确: {months}")
        elif y == min(years.keys()) and len(cols) >= 1:
            pass  # 首年可不足4季度
        else:
            errors.append(f"{y}年只有{len(cols)}个季度，应有4个")
    for c, d in dates:
        cell = ws.cell(1, c)
        if cell.number_format != "YYYY-MM-DD":
            errors.append(f"Col{c} 日期格式={cell.number_format}")
        if not cell.font.bold:
            warnings.append(f"Col{c} 日期未加粗")
        if cell.font.name != "Arial":
            warnings.append(f"Col{c} 日期字体={cell.font.name}")

    # ---- 2. 数字格式 ----
    percent_rows = set()
    amount_rows = {2}
    volume_rows = set()
    for r in range(2, ws.max_row + 1):
        label = str(ws.cell(r, 1).value or "")
        if "yoy" in label.lower() or "毛利率" in label or "占比" in label:
            percent_rows.add(r)
        if "（吨）" in label or "（台" in label or "（套" in label or "（万" in label:
            volume_rows.add(r)
        if "（亿）" in label and "按" not in label:
            amount_rows.add(r)

    fmt_errors = 0
    for r in range(2, ws.max_row + 1):
        if r in percent_rows:
            expected = "0.00%"
        elif r in volume_rows:
            expected = "#,##0.00"
        elif r in amount_rows:
            expected = "0.00"
        else:
            continue
        for c in range(2, min(len(dates) + 2, ws.max_column + 1)):
            cell = ws.cell(r, c)
            if cell.value is not None and cell.number_format != expected:
                fmt_errors += 1
    if fmt_errors > 5:
        warnings.append(f"{fmt_errors}个单元格数字格式不符预期")

    # ---- 3. 空值检查 ----
    dash_count = 0
    for r in range(2, ws.max_row + 1):
        for c in range(2, ws.max_column + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and v.strip() in ("—", "--", "－", "——"):
                dash_count += 1
    if dash_count:
        errors.append(f"Sheet 1 发现有 {dash_count} 个横线占位符（应留空）")

    # ---- 4. 交叉验证 ----
    annual_cols = {}
    for c, d in dates:
        if d.month == 12 and d.day == 31:
            annual_cols[d.year] = c

    for year, col in sorted(annual_cols.items()):
        total = ws.cell(2, col).value
        if not total:
            continue
        products = []
        for r in [28, 33, 38, 43, 48, 53, 58, 63, 68, 73]:
            v = ws.cell(r, col).value
            if v and isinstance(v, (int, float)):
                products.append(v)
        if products:
            prod_sum = sum(products)
            diff = abs(total - prod_sum)
            if diff > 0.05:
                warnings.append(f"{year}年 产品合计={prod_sum:.4f} vs 总收入={total:.4f} diff={diff:.4f}")

        exp = ws.cell(80, col).value
        dom = ws.cell(85, col).value
        if exp and dom:
            region_sum = exp + dom
            diff = abs(total - region_sum)
            if diff > 0.1:
                warnings.append(f"{year}年 地区合计={region_sum:.4f} vs 总收入={total:.4f} diff={diff:.4f}")

    # ---- 输出 ----
    if errors:
        print("❌ 错误：")
        for e in errors:
            print(f"  {e}")
    if warnings:
        print("⚠️ 警告：")
        for w in warnings:
            print(f"  {w}")
    if not errors and not warnings:
        print("✅ 全部通过")

    return len(errors) == 0

if __name__ == "__main__":
    path = sys.argv[1] if len(sys.argv) > 1 else "sample_outputs.xlsx"
    ok = verify(path)
    sys.exit(0 if ok else 1)