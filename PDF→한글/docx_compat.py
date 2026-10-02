"""Normalize pdf2docx output for HWP import without replacing document content."""
from statistics import median
import re
from docx import Document
from docx.oxml.ns import qn


def repair_grid(table):
    """Recover shared grid boundaries from measured cell widths, including spans.

    pdf2docx writes measured tcW values but leaves tblGrid equally divided.
    Only repair when every boundary is constrained and rows agree.
    """
    grid = table.find(qn('w:tblGrid'))
    if grid is None:
        return False
    columns = grid.findall(qn('w:gridCol'))
    samples = {0: [0]}
    vertical_merges = {}
    continued_widths = []
    for row in table.findall(qn('w:tr')):
        if row.find(qn('w:trPr') + '/' + qn('w:gridBefore')) is not None:
            return False
        index, position = 0, 0
        for cell in row.findall(qn('w:tc')):
            props = cell.find(qn('w:tcPr'))
            width = props.find(qn('w:tcW')) if props is not None else None
            if width is None or width.get(qn('w:type')) != 'dxa':
                return False
            span = props.find(qn('w:gridSpan'))
            count = int(span.get(qn('w:val'))) if span is not None else 1
            key = (index, count)
            measured = int(width.get(qn('w:w')))
            merge = props.find(qn('w:vMerge'))
            if merge is not None and merge.get(qn('w:val'), 'continue') == 'continue':
                if key not in vertical_merges:
                    return False
                measured = vertical_merges[key]
                continued_widths.append((width, measured))
            elif merge is not None:
                vertical_merges[key] = measured
            else:
                vertical_merges.pop(key, None)
            index += count
            position += measured
            samples.setdefault(index, []).append(position)
        if index != len(columns):
            return False
    if set(samples) != set(range(len(columns) + 1)):
        return False
    if any(max(values) - min(values) > 4 for values in samples.values()):
        return False
    boundaries = [round(median(samples[i])) for i in range(len(columns) + 1)]
    widths = [b-a for a,b in zip(boundaries, boundaries[1:])]
    if any(width <= 0 for width in widths):
        return False
    for col, width in zip(columns, widths):
        col.set(qn('w:w'), str(width))
    for cell_width, value in continued_widths:
        cell_width.set(qn('w:w'), str(value))
    preferred = table.find(qn('w:tblPr') + '/' + qn('w:tblW'))
    if preferred is not None:
        preferred.set(qn('w:type'), 'dxa')
        preferred.set(qn('w:w'), str(boundaries[-1]))
    return True


def normalize_docx(path):
    doc = Document(path)
    stats = {'tables_repaired': 0, 'tables_unresolved': 0, 'border_sides_normalized': 0}
    aliases = {'MalgunGothic': '맑은 고딕', 'MalgunGothicBold': '맑은 고딕',
               'Batang': '바탕', 'Gulim': '굴림', 'Dotum': '돋움',
               'Gungsuh': '궁서', 'Gothic': '맑은 고딕',
               'GlyphLessFont': '맑은 고딕', 'Heiti': '맑은 고딕'}
    # Keep original families; map only known PDF/PostScript aliases.
    for fonts in doc.element.xpath('.//w:rFonts'):
        for attr in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
            name = fonts.get(qn('w:' + attr))
            if name:
                name = re.sub(r'^[A-Z]{6}\+', '', name)
                fonts.set(qn('w:' + attr), aliases.get(name, name))
    for table in doc.element.xpath('.//w:tbl'):
        repaired = repair_grid(table)
        stats['tables_repaired' if repaired else 'tables_unresolved'] += 1
        # pdf2docx emits LTR tables. Old physical side names import in HWP.
        for parent in table.xpath('./w:tblPr/w:tblBorders | ./w:tblPr/w:tblCellMar | ./w:tr/w:tc/w:tcPr/w:tcBorders | ./w:tr/w:tc/w:tcPr/w:tcMar'):
            for child in list(parent):
                local = child.tag.rsplit('}', 1)[-1]
                if local in ('start', 'end'):
                    child.tag = qn('w:' + {'start': 'left', 'end': 'right'}[local])
                    stats['border_sides_normalized'] += 1
                size = child.get(qn('w:sz'))
                if size is not None:
                    child.set(qn('w:sz'), str(max(2, min(96, round(float(size))))))
                color = child.get(qn('w:color'))
                if color and color.startswith('#'):
                    child.set(qn('w:color'), color[1:])
        # Merged cells may contain duplicate margin elements, invalid in OOXML.
        for props in table.xpath('./w:tr/w:tc/w:tcPr'):
            margins = props.findall(qn('w:tcMar'))
            for extra in margins[1:]:
                for side in extra:
                    old = margins[0].find(side.tag)
                    if old is not None:
                        margins[0].remove(old)
                    margins[0].append(side)
                props.remove(extra)
        indent = table.find(qn('w:tblPr') + '/' + qn('w:tblInd'))
        if indent is not None:
            indent.set(qn('w:w'), str(round(float(indent.get(qn('w:w'), '0')))))
    doc.save(path)
    return stats
