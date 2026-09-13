#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成空白拆分模板（A列标签骨架 + 四行指标块），供换公司时复用。

用法：python scripts/make_blank_template.py templates/空白拆分模板.xlsx
"""
import sys
from openpyxl import Workbook
from openpyxl.styles import Font

DATES = ["2024-06-30", "2024-12-31", "2025-03-31", "2025-06-30",
         "2025-09-30", "2025-12-31", "2026-03-31", "2026-06-30"]

# 每块：(板块标题, [项目名占位...])；每个项目按 名称/yoy/毛利率/占比 四行一组
BLOCKS = [
    ("总收入", ["总收入"]),
    ("按行业（亿）", ["示例行业A", "示例行业B"]),
    ("按产品（亿）", ["示例产品A", "示例产品B", "其他"]),
    ("按地区（亿）", ["内销", "外销"]),
    ("主要产品销量", ["示例产品A（单位）"]),
    ("前五大客户", ["前五大客户销售额（亿）"]),
    ("前五大供应商", ["前五大供应商采购额（亿）"]),
]


def build(path):
    wb = Workbook()
    ws = wb.active
    ws.title = "预测模型"
    ws.cell(1, 1, "项目").font = Font(name="Arial", size=11, bold=True)
    for i, d in enumerate(DATES):
        c = ws.cell(1, 2 + i, d)
        c.font = Font(name="Arial", size=11, bold=True)
    ws.freeze_panes = "B2"

    r = 2
    for title, items in BLOCKS:
        ws.cell(r, 1, title)
        r += 1
        for name in items:
            ws.cell(r, 1, name)
            ws.cell(r + 1, 1, "  yoy")
            ws.cell(r + 2, 1, "  毛利率")
            ws.cell(r + 3, 1, "  占比")
            r += 4
        r += 2  # 板块之间空 2 行：留空不写占位符（Sheet 1 空值一律 None）
    wb.save(path)
    return path


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else "templates/空白拆分模板.xlsx"
    p = build(out)
    # 自检：日期行 + 块数写对了才算生成成功
    from openpyxl import load_workbook
    ws = load_workbook(p).active
    assert ws.cell(1, 1).value == "项目", "表头缺失"
    assert ws.cell(1, 2).value == DATES[0], "日期行缺失"
    labels = [ws.cell(i, 1).value for i in range(1, ws.max_row + 1)]
    assert "总收入" in labels and "按行业（亿）" in labels and "前五大供应商" in labels, "板块缺失"
    print(f"OK -> {p} ({ws.max_row} 行, {ws.max_column} 列)")
