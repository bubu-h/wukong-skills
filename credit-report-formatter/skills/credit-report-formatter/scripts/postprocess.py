import xml.etree.ElementTree as ET
from zipfile import ZipFile
import shutil
import os

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

def fix_docx(input_path, output_path):
    tmp_dir = input_path + '.extracted'
    os.makedirs(tmp_dir, exist_ok=True)

    # Extract
    with ZipFile(input_path, 'r') as z:
        z.extractall(tmp_dir)

    doc_path = os.path.join(tmp_dir, 'word', 'document.xml')
    tree = ET.parse(doc_path)
    root = tree.getroot()

    # Fix tables: full borders + center alignment + cell content format
    for tbl in root.iter(f'{W}tbl'):
        tblPr = tbl.find(f'{W}tblPr')
        if tblPr is None:
            tblPr = ET.Element(f'{W}tblPr')
            tbl.insert(0, tblPr)

        # Table overall center alignment
        jc = tblPr.find(f'{W}jc')
        if jc is None:
            jc = ET.Element(f'{W}jc')
            tblPr.append(jc)
        jc.set(f'{W}val', 'center')

        # Full borders
        borders = tblPr.find(f'{W}tblBorders')
        if borders is None:
            borders = ET.Element(f'{W}tblBorders')
            tblPr.append(borders)
        # Clear existing borders
        for child in list(borders):
            borders.remove(child)
        for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
            b = ET.Element(f'{W}{border_name}')
            b.set(f'{W}val', 'single')
            b.set(f'{W}sz', '4')
            b.set(f'{W}space', '0')
            b.set(f'{W}color', 'auto')
            borders.append(b)

        # Cell content: center alignment, no indent, no spacing, single line height
        for tr in tbl.iter(f'{W}tr'):
            for tc in tr.iter(f'{W}tc'):
                for p in tc.iter(f'{W}p'):
                    cppr = p.find(f'{W}pPr')
                    if cppr is None:
                        cppr = ET.Element(f'{W}pPr')
                        p.insert(0, cppr)

                    # Center alignment
                    cjc = cppr.find(f'{W}jc')
                    if cjc is None:
                        cjc = ET.Element(f'{W}jc')
                        cppr.append(cjc)
                    cjc.set(f'{W}val', 'center')

                    # Remove indent (firstLine=0)
                    cind = cppr.find(f'{W}ind')
                    if cind is None:
                        cind = ET.Element(f'{W}ind')
                        cppr.append(cind)
                    if f'{W}firstLine' in cind.attrib:
                        del cind.attrib[f'{W}firstLine']
                    cind.set(f'{W}firstLine', '0')

                    # Spacing: before=0, after=0, single line height
                    csp = cppr.find(f'{W}spacing')
                    if csp is None:
                        csp = ET.Element(f'{W}spacing')
                        cppr.append(csp)
                    csp.set(f'{W}before', '0')
                    csp.set(f'{W}after', '0')
                    csp.set(f'{W}line', '240')
                    csp.set(f'{W}lineRule', 'auto')

    # Write back
    tree.write(doc_path, encoding='UTF-8', xml_declaration=True)

    # Repack
    with ZipFile(output_path, 'w') as zout:
        for root_dir, dirs, files in os.walk(tmp_dir):
            for file in files:
                file_path = os.path.join(root_dir, file)
                arcname = os.path.relpath(file_path, tmp_dir)
                zout.write(file_path, arcname)

    # Cleanup
    shutil.rmtree(tmp_dir)
    print(f'Post-processed: {output_path}')

if __name__ == '__main__':
    fix_docx(r'd:\code\demo-output.docx', r'd:\code\demo-output-v2.docx')
