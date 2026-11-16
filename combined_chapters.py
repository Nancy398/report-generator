"""Combine single-property 1438 reports by CHAPTER, keeping property metrics separate."""
import io
import re
from copy import deepcopy
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
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
    p.paragraph_format.space_before = Pt(15)
    p.paragraph_format.space_after = Pt(7)
    run = p.add_run(f'{index:02d}  {name}')
    run.bold = True
    run.font.size = Pt(12)
    run.font.name = 'Arial'
    run.font.color.rgb = RGBColor(36, 77, 107)
    return p

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
            out.add_page_break()
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
                if el.tag != qn('w:sectPr'):
                    body.insert(len(body)-1, deepcopy(el))
    result = io.BytesIO()
    out.save(result)
    return result.getvalue()
