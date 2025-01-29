from pdf2docx import parse
import pandas as pd


for i in range(100):
    print('Converting file_{}'.format(i+1))
    # pdf doc
    pdf_document = "D:/WechatAI/WebContent/PDF/file_{}.pdf".format(i+1)
    # word doc
    output_word = "D:/WechatAI/WebContent/Word/file_{}.docx".format(i+1)
    parse(pdf_document, output_word)

print('Finish Converting')
