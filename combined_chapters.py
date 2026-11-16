"""Combine single-property 1438 reports by CHAPTER, keeping property metrics separate."""
import io
import re
from copy import deepcopy
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_BREAK
from reporting import make_docx

CHAPTERS = ('01', '02', '03', '04', '05')

def _text(element):
    return ''.join(n.text or '' for n in element.iter(qn('w:t')))

def _chapter_index(elements):
    found = {}
    for i, el in enumerate(elements):
        if el.tag == qn('w:p'):
            match = re.match(r'^\s*(0[1-5])\s{2,}', _text(el))
            if match:
                found[match.group(1)] = i
    if set(found) != set(CHAPTERS):
        raise ValueError('Single-property 1438 template chapter structure has changed; cannot merge safely.')
    return found

def _add_property_label(doc, name, index, chinese):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(f'{index:02d}  {name}')
    run.bold = True
    run.font.size = Pt(10)
    run.font.name = 'Arial'
    run.font.color.rgb = RGBColor(36, 77, 107)
    return p

def _is_page_break_only(el):
    """The single-property template has an explicit page break before every section.
    Never copy those into a chapter-combined document.
    """
    if el.tag != qn('w:p') or _text(el).strip():
        return False
    return bool(el.xpath('.//w:br[@w:type="page"]') or el.xpath('./w:pPr/w:pageBreakBefore'))


def _compact(el):
    """Reduce inherited oversized spacing, without altering fonts or tables."""
    if el.tag != qn('w:p'):
        return el
    ppr = el.find(qn('w:pPr'))
    if ppr is None:
        from docx.oxml import OxmlElement
        ppr=OxmlElement('w:pPr'); el.insert(0,ppr)
    spacing=ppr.find(qn('w:spacing'))
    if spacing is None:
        from docx.oxml import OxmlElement
        spacing=OxmlElement('w:spacing');ppr.append(spacing)
    spacing.set(qn('w:before'),'0')
    spacing.set(qn('w:after'),'45')
    return el


def make_chapter_combined_docx(llc, month, entries, chinese=False):
    """One cover; each of 5 report chapters contains all properties in sequence.

    entries: [{name, row, property, previous(optional)}]. Each row has its OWN
    occupancy, income, cash, reserve, and distribution. No LLC-level summation.
    """
    if len(entries) < 2:
        raise ValueError('Combined report needs at least two properties.')
    source_docs = []
    for item in entries:
        row = dict(item['row'])
        row['Month (YYYY-MM)'] = month
        prop = dict(item['property'])
        prop['Property Address'] = item['name']
        # Property name should appear in every section and its notes.
        prop['Property Name / LLC'] = item['name']
        doc_bytes = make_docx(row, prop, chinese, previous=item.get('previous'),
                              appendix_documents=[('AppFolio Income Statement', month),
                                                  ('AppFolio Rent Roll', month)])
        doc = Document(io.BytesIO(doc_bytes))
        elements = list(doc.element.body)
        source_docs.append((elements, _chapter_index(elements)))
    # Start with the first report's exact original 1438 cover (and its logo).
    out = Document(io.BytesIO(make_docx(
        dict(entries[0]['row'], **{'Month (YYYY-MM)':month}),
        {'Property Name / LLC':llc,
         'Property Address':f'{len(entries)} Properties  |  Combined LLC Report'},
        chinese, appendix_documents=[])))
    body = out.element.body
    first_elements, first_indices = source_docs[0]
    for el in list(body):
        if el.tag != qn('w:sectPr'):
            body.remove(el)
    # Cover stops just before the first running banner.
    for el in first_elements[:first_indices['01'] - 1]:
        if not _is_page_break_only(el):
            body.insert(len(body) - 1, deepcopy(el))
    # Replace first-report cover fields with LLC name; preserve cover typography.
    cover = [el for el in body if el.tag == qn('w:p')]
    for paragraph_index, value in ((4, llc.upper()), (5, f'{len(entries)} Properties  |  Combined LLC Report')):
        if len(cover) > paragraph_index:
            nodes = list(cover[paragraph_index].iter(qn('w:t')))
            if nodes:
                nodes[0].text = value
                for node in nodes[1:]: node.text = ''
    for c_index, chapter in enumerate(CHAPTERS):
        if c_index:
            out.add_page_break()  # exactly one page break between chapters
        elements, idx = source_docs[0]
        # Single running banner and chapter title, exactly as in 1438.
        body.insert(len(body) - 1, deepcopy(elements[idx[chapter] - 1]))
        body.insert(len(body) - 1, deepcopy(elements[idx[chapter]]))
        for i, item in enumerate(entries):
            elements, indices = source_docs[i]
            start = indices[chapter] + 1
            if c_index < len(CHAPTERS)-1:
                end = indices[CHAPTERS[c_index+1]] - 1  # exclude next banner
            else:
                end = next((j for j in range(start, len(elements)) if elements[j].tag == qn('w:sectPr')),len(elements))
            _add_property_label(out, item['name'], i+1, chinese)
            for el in elements[start:end]:
                if el.tag == qn('w:sectPr') or _is_page_break_only(el):
                    continue
                # Empty template spacer paragraphs take up substantial vertical space.
                if el.tag == qn('w:p') and not _text(el).strip():
                    continue
                body.insert(len(body)-1, _compact(deepcopy(el)))
    # Combined chapters must fit both properties on a single page per section.
    # The original single-property master uses large margins and large table spacing.
    from docx.shared import Inches
    for section in out.sections:
        section.top_margin=Inches(.55)
        section.bottom_margin=Inches(.5)
        section.header_distance=Inches(.22)
        section.footer_distance=Inches(.22)
    for paragraph in out.paragraphs:
        paragraph.paragraph_format.space_before=Pt(0)
        paragraph.paragraph_format.space_after=Pt(2)
        for run in paragraph.runs:
            if run.font.size and run.font.size.pt > 11:
                # Keep cover prominent; compact the rest.
                if paragraph.style.name not in ('Title','Subtitle'):
                    run.font.size=Pt(10)
    for table in out.tables:
        for row in table.rows:
            for cell in row.cells:
                cell.vertical_alignment=1
                for para in cell.paragraphs:
                    para.paragraph_format.space_before=Pt(0)
                    para.paragraph_format.space_after=Pt(0)
                    para.paragraph_format.line_spacing=1.0
                    for run in para.runs:
                        run.font.size=Pt(8)
    result = io.BytesIO()
    out.save(result)
    return result.getvalue()
