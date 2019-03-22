import docx
from docx.document import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import _Cell, Table
from docx.text.paragraph  import Paragraph


def iter_block_items(parent):
    """
    Yield each paragraph and table child within *parent*, in document order.
    Each returned value is an instance of either Table or Paragraph.
    """
    if isinstance(parent, Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("something's not right")

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def docx_to_list(docx):
    result = []
    for block in iter_block_items(docx):
        if isinstance(block, Paragraph):
            result.append(block.text)
        elif isinstance(block, Table):
            temp = []
            for row in block.rows:
                cell_list = []
                for cell in row.cells:
                    for elment in iter_block_items(cell):
                        cell_list.append(elment.text)
                temp.append(cell_list)
            result.append(temp)
    return result


def print_docx(docx):
    for block in iter_block_items(docx):
        if isinstance(block, Paragraph):
            print(block.text)
        elif isinstance(block, Table):
            for row in block.rows:
                for cell in row.cells:
                    for elment in iter_block_items(cell):
                        print(elment.text+'\t', end="")
                print('\n')


def delete_paragraph(paragraph):
    p = paragraph._element
    p.getparent().remove(p)
    p._p = p._element = None