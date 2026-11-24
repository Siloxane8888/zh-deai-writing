# zh-deai-writing

降低中文写作里的 **AI 痕迹**，同时**不动文档格式**。

给一份 `.docx` 论文或报告，AI 只回来一份"哪一段改成什么"的补丁清单，由脚本就地替换文字。标题样式、目录、自动编号、页眉页脚、表格、图片、引用标记，全部保持原样。

> 起因很实际：把 Word 论文直接丢给 AI 润色，它会回你一份纯文本全文；你再存成 Word，标题样式没了、目录变成死字、页码全部重来。

---

## 它解决的两个问题

**一、改哪里**

AI 味的主要载体不是词表，是**均匀度**——段落长度、句长、段落开头方式、并列结构，四项都过于整齐。刻意避开"维度""赋能""值得注意的是"之后，文本照样能被认出。

真实测过的一篇 4200 字中文课程报告：

| | 机器生成的 23 段 | 人工插入的 5 段 |
|---|---|---|
| 段长均值 / 中位 | 151.4 / 147.0 | 119.0 / 99.0 |
| 段长变异系数 CV | **0.19** | **0.70** |
| 长度区间 | 106–222 | 33–270 |

均值和中位几乎重合，说明分布对称没有长尾——批量生成的典型形状。同一篇里还有四处连续 32 字的句子、三处连续 38 字的句子，人写不出来。

详见 [docs/why-uniformity.md](docs/why-uniformity.md)。

**二、怎么改而不毁格式**

`.docx` 不是文本文档，是压缩包里的关系图。正文 XML 里一个段落由 `pPr`（样式/缩进/间距/编号）+ 若干 `rPr`（字体/字号/加粗）+ `w:t`（文字）组成。AI 回传全文，等于把前两者全丢了。

所以这里的做法是**补丁模式**：AI 输出 `{anchor, new_text}`，脚本找到对应段落，保留 `pPr` 和第一个 run 的 `rPr`，只换文字。

详见 [docs/docx-safety.md](docs/docx-safety.md)。

---

## 快速开始

```bash
pip install python-docx
```

**1. 拍结构快照**（改之前）

```bash
python scripts/docx_fingerprint.py 论文.docx
```

会打印段落数、样式分布、表格、图片、域代码、页眉页脚，以及**段内格式混排的段落**——这些段落不能整段改写，先看清楚。

**2. 用提示词**

复制 [prompt/zh-deai.prompt.md](prompt/zh-deai.prompt.md) 里代码块的内容，末尾粘上正文或上传 docx。AI 会回一个 JSON。

存成 `edits.json`。

**3. 空跑**

```bash
python scripts/apply_edits.py 论文.docx edits.json --dry-run
```

检查 `applied` / `skipped` / `ambiguous` / `not_found`。有 ambiguous（anchor 不唯一）就让它写长一点；有 not_found 就是 anchor 没对上。

**4. 应用**

```bash
python scripts/apply_edits.py 论文.docx edits.json --out 论文-改.docx
```

**5. 验结构**

```bash
python scripts/docx_fingerprint.py 论文.docx 论文-改.docx
```

必须输出「结构一致」。不通过就别交。

**6. 收尾**

Word 里 `Ctrl+A` → `F9` 更新目录页码。封面的姓名学号自己填。

---

## 文件

```
prompt/
  zh-deai.prompt.md    主提示词，docx 安全版（推荐）
  zh-deai.lite.md      精简版，篇幅不够时用
scripts/
  apply_edits.py       按 JSON 补丁改文字，不碰格式
  docx_fingerprint.py  结构指纹与比对
docs/
  why-uniformity.md    判断依据：为什么均匀度是核心指纹
  docx-safety.md       为什么会毁格式，以及安全替换的机制
examples/
  edits.example.json   补丁格式示例
```

---

## 脚本会主动拒绝的情况

`apply_edits.py` 只在完全安全时才改。以下一律跳过并在报告里说明：

| 段落含 | 原因 |
|---|---|
| 域代码（`fldChar` / `instrText`） | 目录、页码、表图交叉引用会断 |
| 图片 / 脚注 / 批注 | 重建 run 会丢锚点 |
| 超链接 | 链接文字不在 `p.runs` 里，硬改会重复 |
| 手打换行 / 制表位 | 会被压平 |
| anchor 命中多个段落 | 定位不唯一，宁可不动 |

表格内文字默认锁定，要改得显式加 `--allow-table`。

---

## 适用范围

- 用于**自己写的**中文稿子：论文、课程报告、技术文档、工作汇报。
- 提示词里有一条硬约束：不编造原文没有的数字、经历、课程名和引用。降低痕迹的正确方向是**把作者放回去**，不是给机器文本加装饰。
- 元数据（累计编辑时长、修订次数）和真实记录（哪道题、什么报错）管不了。这两样只能靠真的动手。
- 保留并写准 AI 使用声明。读者判断"是不是 AI 写的"，看的是作者有没有自己的东西，不是有没有用过工具。

---

## License

MIT
