import re,os
from docx import Document
from docx.shared import Inches


def convert(path):
    if path == '':
        path = 'script.docx'
    elif not path.endswith('.docx'):
        path += '.docx'
    if path not in os.listdir():
        print('That file does not exist.')
        return None
    read = Document(path)
    write = Document()
    for paragraph in read.paragraphs:
        paras = re.split('<|>',paragraph.text)
        if len(paras) == 1:
            format = write.add_paragraph(paras[0])
        elif len(paras) > 2:
            paras = paras[1:]
            header = paras[0].strip().upper()
            if header == 'INT' or header == 'EXT' or header == 'SUB':
                format = write.add_paragraph(header+'. '+paras[1].strip().upper())
            else:

                head_f = write.add_paragraph(header)
                head_f.paragraph_format.left_indent = Inches(2.0)
                body_f = write.add_paragraph(paras[1].strip())
                body_f.paragraph_format.left_indent = Inches(1.0)


    write.save('formated'+path)




def driver():
    path = ''
    #docx = DOCX()
    while path != 'exit':
        prompt = 'Enter document to convert. Enter "exit" to end. "script.docx" is default.\n'
        path = input(prompt)
        path = path.strip().lower()
        if path == 'exit':
            return print('Program ended')
        convert(path)


driver()

