# coding=utf-8

# 此程序用于转换word的格式

from win32com import client


# doc格式转docx格式，docx格式是基于xml的，doc格式不是基于xml的，docx库只能解析docx格式的文件
def doc2docx(doc_name, docx_name):
    word = client.Dispatch("Word.Application")
    # 是否显示，默认是False，不显示
    # word.Visible = True
    # 是否显示警告，默认是False，不显示
    # word.DisplayAlerts = True
    doc = word.Documents.Open(doc_name)
    # 使用参数16表示将doc转换成docx
    doc.SaveAs(docx_name, 16)
    doc.Close()
    word.Quit()


def docx2doc(doc_name, docx_name):
    word = client.Dispatch("Word.Application")
    doc = word.Documents.Open(docx_name)
    # 使用参数0表示将docx转换成doc
    doc.SaveAs(doc_name, 0)
    doc.Close()
    word.Quit()


def doc2html(doc_name, html_name):
    word = client.Dispatch("Word.Application")
    doc = word.Documents.Open(doc_name)
    # 使用参数8表示将doc转换成html
    doc.SaveAs(html_name, 8)
    doc.Close()
    word.Quit()