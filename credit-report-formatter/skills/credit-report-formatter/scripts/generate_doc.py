#!/usr/bin/env python3
"""
授信报告格式调整工具
用法: python generate_doc.py <input.md> [-o <output.docx>]
"""

import sys
import os
import subprocess
import tempfile
import shutil
from pathlib import Path
from zipfile import ZipFile

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def get_skill_dir():
    return Path(__file__).parent.parent


def detect_platform():
    p = sys.platform
    if p == "win32":
        return "windows"
    elif p == "darwin":
        return "macos"
    else:
        raise RuntimeError(f"不支持的平台: {p}")


def postprocess_docx(input_path: str, output_path: str):
    """对 pandoc 生成的 docx 做后处理: 修表格、修 Heading 缩进"""
    tmp_dir = input_path + ".extracted"
    os.makedirs(tmp_dir, exist_ok=True)

    with ZipFile(input_path, "r") as z:
        z.extractall(tmp_dir)

    doc_path = os.path.join(tmp_dir, "word", "document.xml")

    try:
        import xml.etree.ElementTree as ET

        ET.register_namespace(
            "w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        )

        tree = ET.parse(doc_path)
        root = tree.getroot()

        # 修表格（模板样式无法继承表格边框和居中，必须后处理）
        for tbl in root.iter(f"{W}tbl"):
            tblPr = tbl.find(f"{W}tblPr")
            if tblPr is None:
                tblPr = ET.Element(f"{W}tblPr")
                tbl.insert(0, tblPr)

            # 表格整体居中
            jc = tblPr.find(f"{W}jc")
            if jc is None:
                jc = ET.Element(f"{W}jc")
                tblPr.append(jc)
            jc.set(f"{W}val", "center")

            # 全边框
            borders = tblPr.find(f"{W}tblBorders")
            if borders is None:
                borders = ET.Element(f"{W}tblBorders")
                tblPr.append(borders)
            for child in list(borders):
                borders.remove(child)
            for name in ["top", "left", "bottom", "right", "insideH", "insideV"]:
                b = ET.Element(f"{W}{name}")
                b.set(f"{W}val", "single")
                b.set(f"{W}sz", "4")
                b.set(f"{W}space", "0")
                b.set(f"{W}color", "auto")
                borders.append(b)

            # 单元格内容: 居中、无缩进、段前段后0、单倍行距
            for tr in tbl.iter(f"{W}tr"):
                for tc in tr.iter(f"{W}tc"):
                    for p in tc.iter(f"{W}p"):
                        cppr = p.find(f"{W}pPr")
                        if cppr is None:
                            cppr = ET.Element(f"{W}pPr")
                            p.insert(0, cppr)

                        cjc = cppr.find(f"{W}jc")
                        if cjc is None:
                            cjc = ET.Element(f"{W}jc")
                            cppr.append(cjc)
                        cjc.set(f"{W}val", "center")

                        cind = cppr.find(f"{W}ind")
                        if cind is None:
                            cind = ET.Element(f"{W}ind")
                            cppr.append(cind)
                        if f"{W}firstLine" in cind.attrib:
                            del cind.attrib[f"{W}firstLine"]
                        cind.set(f"{W}firstLine", "0")

                        csp = cppr.find(f"{W}spacing")
                        if csp is None:
                            csp = ET.Element(f"{W}spacing")
                            cppr.append(csp)
                        csp.set(f"{W}before", "0")
                        csp.set(f"{W}after", "0")
                        csp.set(f"{W}line", "240")
                        csp.set(f"{W}lineRule", "auto")

        tree.write(doc_path, encoding="UTF-8", xml_declaration=True)
    except Exception as e:
        print(f"后处理警告: {e}")

    # 重新打包
    with ZipFile(output_path, "w") as zout:
        for root_dir, dirs, files in os.walk(tmp_dir):
            for file in files:
                file_path = os.path.join(root_dir, file)
                arcname = os.path.relpath(file_path, tmp_dir)
                zout.write(file_path, arcname)

    shutil.rmtree(tmp_dir)
    print(f"后处理完成: {output_path}")


def main():
    if len(sys.argv) < 2:
        print("用法: python generate_doc.py <input.md> [-o <output.docx>]")
        sys.exit(1)

    input_md = Path(sys.argv[1]).resolve()
    if not input_md.exists():
        print(f"错误: 找不到文件 {input_md}")
        sys.exit(1)

    if "-o" in sys.argv:
        idx = sys.argv.index("-o")
        output_docx = Path(sys.argv[idx + 1]).resolve() if idx + 1 < len(sys.argv) else None
    else:
        output_docx = input_md.with_suffix(".docx")

    if output_docx is None:
        print("错误: -o 参数后需要跟输出路径")
        sys.exit(1)

    platform = detect_platform()
    skill_dir = get_skill_dir()
    template = skill_dir / "templates" / f"reference-doc-{platform}.docx"

    if not template.exists():
        print(f"错误: 找不到模板 {template}")
        sys.exit(1)

    # 先生成到临时文件
    tmp_docx = tempfile.mktemp(suffix=".docx")
    cmd = [
        "pandoc",
        str(input_md),
        "-o",
        tmp_docx,
        f"--reference-doc={template}",
    ]
    print(f"[{platform}] pandoc: {' '.join(cmd)}")
    subprocess.run(cmd, check=True)

    # 后处理
    postprocess_docx(tmp_docx, str(output_docx))
    os.remove(tmp_docx)

    print(f"完成: {output_docx}")


if __name__ == "__main__":
    main()
