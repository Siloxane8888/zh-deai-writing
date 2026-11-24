#!/usr/bin/env python3
"""比对两个 .docx 的结构指纹，确认"只改了文字，没动格式"。

用法：
    python docx_fingerprint.py 原稿.docx                 # 打印结构指纹与风险段落
    python docx_fingerprint.py 原稿.docx 改后.docx        # 比对，结构不一致时退出码 1

检查项：
    - 段落数、每段段落样式名与段落属性（对齐/缩进/间距/编号）指纹
    - 表格数量与每表行列数
    - 图片数量
    - 域代码数量（目录、页码、表图交叉引用）
    - 节数量与页眉页脚文字
    - 超链接数量
    - 段落内字符格式是否混排（混排段落不适合整段改写）
"""

import hashlib
import sys

from docx import Document

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")


def _h(s):
    return hashlib.sha1(s.encode("utf-8")).hexdigest()[:12]


def _rpr_sig(run):
    """字符格式签名：只看会改变视觉效果的属性，忽略 rFonts 之类的排版提示。"""
    f = run.font
    color = ""
    try:
        if f.color is not None and f.color.rgb is not None:
            color = str(f.color.rgb)
    except (AttributeError, ValueError):
        color = ""
    size = ""
    try:
        if f.size is not None:
            size = str(f.size.pt)
    except (AttributeError, ValueError):
        size = ""
    return "|".join([
        str(f.bold), str(f.italic), str(f.underline), str(f.strike),
        size, color, str(f.name or ""),
    ])


def _para_fp(p):
    ppr = p._p.pPr
    ppr_xml = str(ppr.xml) if ppr is not None else ""
    rprs = set()
    for r in p.runs:
        rprs.add(_rpr_sig(r))
    return {
        "style": p.style.name,
        "ppr": _h(ppr_xml),
        "rpr_variants": len(rprs),
        "text_len": len(p.text),
        "text_head": p.text[:24],
    }


def fingerprint(path):
    doc = Document(path)
    xml = doc.element.xml

    fps = [_para_fp(p) for p in doc.paragraphs]

    sections = []
    for i, sec in enumerate(doc.sections):
        hdr = " / ".join(p.text for p in sec.header.paragraphs if p.text.strip())
        ftr = " / ".join(p.text for p in sec.footer.paragraphs if p.text.strip())
        sections.append({"index": i, "header": hdr, "footer": ftr})

    return {
        "para_count": len(fps),
        "paragraphs": fps,
        "style_histogram": _histogram([f["style"] for f in fps]),
        "tables": [[len(t.rows), len(t.columns)] for t in doc.tables],
        "images": len(doc.inline_shapes),
        "fields": xml.count("fldChar"),
        "hyperlinks": xml.count("<w:hyperlink"),
        "sections": sections,
        "mixed_format_paragraphs": [
            {"index": i, "head": f["text_head"], "variants": f["rpr_variants"]}
            for i, f in enumerate(fps) if f["rpr_variants"] > 1
        ],
    }


def _histogram(values):
    out = {}
    for v in values:
        out[v] = out.get(v, 0) + 1
    return dict(sorted(out.items(), key=lambda kv: -kv[1]))


def summarize(fp):
    print(f"段落数        {fp['para_count']}")
    print(f"表格          {fp['tables']}")
    print(f"图片          {fp['images']}")
    print(f"域代码        {fp['fields']}")
    print(f"超链接        {fp['hyperlinks']}")
    print(f"节            {len(fp['sections'])}")
    for s in fp["sections"]:
        print(f"  节{s['index']}  页眉[{s['header']}]  页脚[{s['footer']}]")
    print("段落样式分布:")
    for k, v in fp["style_histogram"].items():
        print(f"  {v:>3}  {k}")
    mixed = fp["mixed_format_paragraphs"]
    print(f"段内字符格式混排的段落  {len(mixed)}")
    for m in mixed[:10]:
        print(f"  #{m['index']}  {m['variants']} 种格式  {m['head']}")


def compare(a, b):
    problems = []

    for key in ("para_count", "tables", "images", "fields", "hyperlinks"):
        if a[key] != b[key]:
            problems.append(f"{key}: {a[key]} -> {b[key]}")

    if a["sections"] != b["sections"]:
        problems.append("页眉页脚或分节发生变化")

    if a["style_histogram"] != b["style_histogram"]:
        problems.append(f"段落样式分布变化: {a['style_histogram']} -> {b['style_histogram']}")

    n = min(len(a["paragraphs"]), len(b["paragraphs"]))
    style_changed, ppr_changed = [], []
    for i in range(n):
        pa, pb = a["paragraphs"][i], b["paragraphs"][i]
        if pa["style"] != pb["style"]:
            style_changed.append((i, pa["style"], pb["style"], pa["text_head"]))
        elif pa["ppr"] != pb["ppr"]:
            ppr_changed.append((i, pa["text_head"]))

    for i, sa, sb, head in style_changed:
        problems.append(f"段落 #{i} 样式被改: {sa} -> {sb}  [{head}]")
    for i, head in ppr_changed[:20]:
        problems.append(f"段落 #{i} 段落属性被改（对齐/缩进/间距/编号）[{head}]")

    changed_text = sum(
        1 for i in range(n)
        if a["paragraphs"][i]["text_len"] != b["paragraphs"][i]["text_len"]
    )

    return problems, changed_text


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 2

    a = fingerprint(sys.argv[1])

    if len(sys.argv) == 2:
        summarize(a)
        return 0

    b = fingerprint(sys.argv[2])
    problems, changed_text = compare(a, b)

    print(f"文字长度发生变化的段落: {changed_text}")
    if problems:
        print("\n结构不一致：")
        for p in problems:
            print("  ✗", p)
        return 1

    print("结构一致：段落数、样式、段落属性、表格、图片、域、页眉页脚、超链接均未变化。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
