#!/usr/bin/env python3
"""从 .docx 合同提取纯文本（保留段落与表格顺序），供 CRC 审核 Skill 使用。

用法：
    python3 extract_docx.py 合同文件.docx [输出.txt]

- 仅依赖标准库（zipfile + xml），无需 python-docx。
- 输出到 stdout 或指定文件；段落间空行，表格按"行 | 单元格"展开。
- 已知限制：不还原自动编号样式（1.1/2.3 等编号若由 Word 样式生成而非手输，
  提取文本中不会出现编号，审核时按"第 N 段"定位即可）。
"""
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def para_text(p):
    parts = []
    for node in p.iter():
        if node.tag == W + "t":
            parts.append(node.text or "")
        elif node.tag == W + "tab":
            parts.append("\t")
        elif node.tag == W + "br":
            parts.append("\n")
    return "".join(parts).strip()


def table_text(tbl):
    lines = []
    for row in tbl.findall(W + "tr"):
        cells = []
        for cell in row.findall(W + "tc"):
            cell_paras = [para_text(p) for p in cell.findall(W + "p")]
            cells.append(" ".join(t for t in cell_paras if t))
        line = " | ".join(cells).strip()
        if line.strip(" |"):
            lines.append("| " + line + " |")
    return lines


def extract(path):
    with zipfile.ZipFile(path) as z:
        xml_bytes = z.read("word/document.xml")
    root = ET.fromstring(xml_bytes)
    body = root.find(W + "body")
    out = []
    for child in body:
        if child.tag == W + "p":
            t = para_text(child)
            out.append(t)
        elif child.tag == W + "tbl":
            out.extend(table_text(child))
            out.append("")
    text = "\n".join(out)
    text = re.sub(r"\n{3,}", "\n\n", text).strip()
    return text


def main():
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    text = extract(sys.argv[1])
    if len(sys.argv) >= 3:
        with open(sys.argv[2], "w", encoding="utf-8") as f:
            f.write(text + "\n")
        print(f"已提取 {len(text)} 字符 → {sys.argv[2]}", file=sys.stderr)
    else:
        print(text)


if __name__ == "__main__":
    main()
