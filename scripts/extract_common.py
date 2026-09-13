# -*- coding: utf-8 -*-
"""通用文本表格解析helpers：把年报/招股书 PDF 转出的 txt 还原成可用的数值表。"""
import re

NUM_RE = re.compile(r'^-?[\d,]+\.?\d*$')
PCT_RE = re.compile(r'^-?[\d,.]+%$')
YI = 1e8
WAN = 1e4


def clean_num(s):
    return s.replace(',', '').replace(' ', '').replace('\u3000', '').strip()


def is_num(s):
    return bool(NUM_RE.match(s.strip().replace(' ', '')))


def is_pct(s):
    return bool(PCT_RE.match(s.strip().replace(' ', '')))


def merge_broken(lines, i):
    """Return (full_number_string, next_index). Handles A/B/C break modes."""
    s = clean_num(lines[i])
    j = i + 1
    n = len(lines)
    while j < n:
        nxt = clean_num(lines[j])
        if not nxt:
            break
        if s.endswith('.'):
            s = s + nxt
            j += 1
            continue
        if '.' not in s:
            # cleaned string has no commas; any numeric continuation is a fragment merge
            if re.match(r'^[\d.,]+$', nxt):
                s = s + nxt
                j += 1
                continue
            break
        else:
            # s is comma-stripped; mode C tail = ddd.d + single digit
            if s.endswith('.'):
                s = s + nxt          # B: ...641,215. + 37
                j += 1
                continue
            if re.match(r'^\d\.\d$', s) is None and re.search(r'\d{3}\.\d$', s) and re.match(r'^\d$', nxt):
                s = s + nxt          # C: 1,042,215,645.3 + 7
                j += 1
                continue
            break
    return s, j


def strip_label(lab):
    lab = lab.replace('\u3000', ' ')
    # drop page headers / noise
    lab = re.sub(r'===PAGE\s*\d*===', '', lab)
    lab = re.sub(r'^.{4,40}?(股份)?有限公司.*?(年年度报告全文|半年度报告全文|招股说明书|$)', '', lab)  # 去掉页眉里的公司全称
    lab = lab.replace('--', '').replace('-', '')
    lab = re.sub(r'\s+', '', lab)
    return lab.strip()


def block_parse(section_text, pct_value=True):
    """Return [(label, [values], [is_pct_flags])]. Non-numeric lines accumulate into label.
    pct_value=True: '12.34%' lines count as values. '--' lines are dropped."""
    lines = [l.strip() for l in section_text.split('\n')]
    out = []
    label = ''
    vals = []
    flags = []
    i = 0
    n = len(lines)
    while i < n:
        line = lines[i]
        if not line or line == '--' or line == '-':
            i += 1
            continue
        if is_pct(line) and pct_value:
            vals.append(float(clean_num(line).rstrip('%')))
            flags.append(True)
            i += 1
            continue
        if is_num(line) or (line.startswith('-') and is_num(line[1:]) and not is_num(line)):
            num, i = merge_broken(lines, i)
            try:
                f = float(clean_num(num))
            except Exception:
                i += 1
                continue
            vals.append(f)
            flags.append(False)
            continue
        if vals:
            out.append((strip_label(label), vals, flags))
            label = ''
            vals = []
            flags = []
        label += line
        i += 1
    if label or vals:
        out.append((strip_label(label), vals, flags))
    return [(l, v, f) for l, v, f in out if l or v]


def load(path):
    return open(path, encoding='utf-8').read()


def section(text, start_anchor, end_anchors):
    idx = text.find(start_anchor)
    if idx < 0:
        return None
    tail = text[idx + len(start_anchor):]
    ends = [tail.find(a) for a in end_anchors if a in tail]
    ends = [e for e in ends if e >= 0]
    cut = min(ends) if ends else len(tail)
    return tail[:cut]