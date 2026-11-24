#!/usr/bin/env python3
"""按 JSON 补丁列表修改 .docx 的正文文字，不碰任何格式。

只做一件事：把指定段落的文字换掉。
段落样式（pPr）、字符格式（rPr）、目录域、页眉页脚、表格结构、图片、编号，
一律不动。段落数也不变——不合并、不拆分。

用法：
    python apply_edits.py 论文.docx edits.json --out 论文-改.docx
    python apply_edits.py 论文.docx edits.json --dry-run

edits.json 格式见 examples/edits.example.json
"""

import argparse
import copy
import json
import sys

from docx import Document

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

# 段落里出现这些元素就跳过：重建 run 会破坏它们
PROTECTED_TAGS = {
    "fldChar", "instrText", "fldSimple",          # 域：目录、页码、表图交叉引用
    "drawing", "pict", "object",                  # 图片与嵌入对象
    "footnoteReference", "endnoteReference",      # 脚注尾注
    "commentReference", "commentRangeStart",      # 批注
    "hyperlink",                                  # 超链接（文本不在 p.runs 里）
    "br", "cr", "tab",                            # 手动换行与制表位
}


def _paragraph_blockers(p):
    """返回该段落不能安全改写的理由列表。"""
    reasons = []
    for child in p._p.iter():
        tag = child.tag
        if "}" in tag:
            tag = tag.split("}", 1)[1]
        if tag in PROTECTED_TAGS and tag not in reasons:
            reasons.append(tag)
    return reasons


def _visual_sig(run):
    """字符格式签名：只看会改变视觉效果的属性。"""
    f = run.font
    size = ""
    try:
        if f.size is not None:
            size = str(f.size.pt)
    except (AttributeError, ValueError):
        pass
    color = ""
    try:
        if f.color is not None and f.color.rgb is not None:
            color = str(f.color.rgb)
    except (AttributeError, ValueError):
        pass
    return "|".join([str(f.bold), str(f.italic), str(f.underline), str(f.strike),
                     size, color, str(f.name or "")])


def _rebuild(p, new_text):
    """保留第一个 run 的字符格式，把整段文字换成 new_text。"""
    runs = list(p.runs)
    if not runs:
        return False, "no_runs"

    rpr = copy.deepcopy(runs[0]._element.rPr)
    for r in runs:
        r._element.getparent().remove(r._element)

    run = p.add_run()
    old = run._element.rPr
    if old is not None:
        run._element.remove(old)
    if rpr is not None:
        run._element.insert(0, copy.deepcopy(rpr))
    run.text = new_text
    return True, "ok"


def _collect(doc):
    """正文段落 + 表格单元格段落，返回 [(scope, paragraph)]。"""
    items = [("body", p) for p in doc.paragraphs]
    for ti, table in enumerate(doc.tables):
        for ri, row in enumerate(table.rows):
            for ci, cell in enumerate(row.cells):
                for p in cell.paragraphs:
                    items.append((f"table[{ti}][{ri}][{ci}]", p))
    return items


def main():
    ap = argparse.ArgumentParser(description="按 JSON 补丁修改 docx 文字，保留全部格式")
    ap.add_argument("docx")
    ap.add_argument("edits", help="JSON 补丁文件")
    ap.add_argument("--out", help="输出文件，默认覆盖输入")
    ap.add_argument("--dry-run", action="store_true", help="只报告匹配结果，不写文件")
    ap.add_argument("--allow-table", action="store_true", help="允许修改表格内段落")
    args = ap.parse_args()

    with open(args.edits, encoding="utf-8-sig") as f:
        spec = json.load(f)

    edits = spec["edits"] if isinstance(spec, dict) else spec
    doc = Document(args.docx)

    report = {"applied": [], "skipped": [], "ambiguous": [], "not_found": []}

    for idx, edit in enumerate(edits):
        anchor = edit.get("anchor", "")
        new_text = edit.get("new_text", "")
        want_scope = edit.get("scope", "body")

        if not anchor or not new_text:
            report["skipped"].append({"index": idx, "anchor": anchor,
                                      "reason": "missing_anchor_or_text"})
            continue

        hits = [(scope, p) for scope, p in _collect(doc) if anchor in p.text]

        # 按 scope 收窄候选：body 只找正文，table* 只找表格
        if want_scope == "body":
            narrowed = [h for h in hits if h[0] == "body"]
        elif want_scope.startswith("table"):
            narrowed = [h for h in hits if h[0].startswith("table")]
        else:
            narrowed = hits
        if narrowed:
            hits = narrowed

        if len(hits) == 0:
            report["not_found"].append({"index": idx, "anchor": anchor})
            continue
        if len(hits) > 1:
            report["ambiguous"].append({
                "index": idx,
                "anchor": anchor,
                "scopes": [s for s, _ in hits],
                "hint": "anchor 不唯一，请写长一点",
            })
            continue

        scope, p = hits[0]
        if scope != "body" and not args.allow_table:
            report["skipped"].append({"index": idx, "anchor": anchor, "scope": scope,
                                      "reason": "table_locked (加 --allow-table 放行)"})
            continue

        blockers = _paragraph_blockers(p)
        if blockers:
            report["skipped"].append({"index": idx, "anchor": anchor, "scope": scope,
                                      "reason": "protected:" + ",".join(blockers)})
            continue

        if len({_visual_sig(r) for r in p.runs}) > 1:
            report["skipped"].append({"index": idx, "anchor": anchor, "scope": scope,
                                      "reason": "mixed_format（段内有加粗/斜体等混排，整段合并会抹掉）"})
            continue

        old_text = p.text
        ok, why = _rebuild(p, new_text)
        entry = {"index": idx, "anchor": anchor, "scope": scope,
                 "old_len": len(old_text), "new_len": len(new_text)}
        if ok:
            report["applied"].append(entry)
        else:
            report["skipped"].append({"index": idx, "anchor": anchor, "reason": why})

    if not args.dry_run:
        out = args.out or args.docx
        doc.save(out)
        report["output"] = out

    print(json.dumps(report, ensure_ascii=False, indent=2))

    return 1 if (report["ambiguous"] or report["not_found"]) else 0


if __name__ == "__main__":
    sys.exit(main())
