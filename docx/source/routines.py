from docx import Document
from docx import shared
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.shared import RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Inches
from pathlib import Path
import win32com.client
import pandas as pd


def read_template(path: str) -> Document:
    """
    read docx template
    """

    path = Path(path)
    if not Path.exists(path):
        msg = f'template_path {path} does not exist!'
        raise ValueError(msg)

    if not Path(path).suffix in ['.docx']:
        msg = 'file type must be ".docx"!'
        raise ValueError(msg)

    document = Document(path)

    return document


def _add_caption(caption):
    """
    add caption to figure
    """

    run = caption.add_run()
    r = run._r
    fldChar = OxmlElement('w:fldChar')
    fldChar.set(qn('w:fldCharType'), 'begin')
    r.append(fldChar)
    instrText = OxmlElement('w:instrText')
    instrText.text = ' SEQ Figure * ARABIC'
    r.append(instrText)
    fldChar = OxmlElement('w:fldChar')
    fldChar.set(qn('w:fldCharType'), 'end')
    r.append(fldChar)


def add_header(document: Document,
               header: str,
               font_size=11) -> Document:
    """
    add header to page
    """

    run = document.add_heading().add_run(header)
    run.alignment = WD_ALIGN_PARAGRAPH.CENTER
    font = run.font
    font.size = shared.Pt(font_size)
    font.color.rgb = RGBColor(0, 0, 0)

    return document


def add_image(document: Document,
              image_path: str,
              caption='Figure 0: Example Data Plot',
              width=5.4,
              height=3.61,
              font_size=9,
              new_page=False) -> Document:
    """
    add image to document
    """

    # format paragraph to be added
    p = document.add_paragraph()
    paragraph_format = p.paragraph_format
    paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run()
    if new_page:
        r.add_break()
    r.add_picture(image_path,
                  width=Inches(width),
                  height=Inches(height))

    # update style and font
    style = document.styles['Caption']
    font = style.font
    font.size = shared.Pt(font_size)
    font.color.rgb = RGBColor(0, 0, 0)
    font.italic = True
    paragraph = document.add_paragraph(caption)
    paragraph_format = paragraph.paragraph_format
    paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # add caption
    _add_caption(paragraph)

    return document


def add_table(document: Document,
              table_in: pd.DataFrame,
              caption='Example Table 0: Summary Data',
              font_size=9,
              table_style='ams_table_style',
              new_page=False) -> Document:
    """
    add table to document
    """

    if new_page:
        document.add_page_break()

    # format text and table to be added
    p = document.add_paragraph()
    paragraph_format = p.paragraph_format
    paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    style = document.styles['Caption']
    font = style.font
    font.size = shared.Pt(font_size)
    font.color.rgb = RGBColor(0, 0, 0)
    font.italic = True
    paragraph = document.add_paragraph(caption, style='Caption')
    paragraph_format = paragraph.paragraph_format
    paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _add_caption(paragraph)

    # add table
    table = document.add_table(table_in.shape[0]+1, table_in.shape[1])

    # add header rows
    for j in range(table_in.shape[-1]):
        table.cell(0, j).text = table_in.columns[j]

    # add the rest of the table
    for i in range(table_in.shape[0]):
        for j in range(table_in.shape[-1]):
            table.cell(i+1, j).text = str(table_in.values[i, j])
            table.cell(i+1, j).vertical_alignment =\
                WD_ALIGN_VERTICAL.CENTER

    # format table in document
    try:
        table.style = table_style
    except:
        msg = f'could not set table_stlye=f{table_style}'
        print(msg)

    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    p2 = document.add_paragraph()
    paragraph_format = p2.paragraph_format
    paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER

    return document


def update_toc(file: str):
    """
    Subroutine for updating TOC after the entire
    document has been built and saved

    Args:
        file (str): Full path to document
    """
    word = win32com.client.DispatchEx("Word.Application")
    doc = word.Documents.Open(file)
    doc.TablesOfContents(1).Update()
    doc.Close(SaveChanges=True)
    word.Quit()


def save(document: Document, path: str):
    """
    """

    document.save(path)


def append_document(main_document, sup_document):
    """
    append a supplemental document to the main document
    """

    for element in sup_document.element.body:
        main_document.element.body.append(element)

    return main_document
