from docx import Document
from docx.shared import Inches
from docx.shared import Pt
import re, os


def open_read(path):
    if path == '':
        path = 'script.docx'
    elif not path.endswith('.docx'):
        path += '.docx'
    if path not in os.listdir():
        return print('That file does not exist.')
    return Document(path), path


def open_write():
    doc = Document()
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Courier New'
    font.size = Pt(12)
    return doc


def description(doc,paragraph):
    format = doc.add_paragraph(paragraph)
    format.style = doc.styles['Normal']


def is_subheader(header):
    return header == 'INT' or header == 'EXT' or header == 'SUB'


def heading(doc,header,body):
    format = doc.add_paragraph(header+'. '+body.strip().upper())
    format.style = doc.styles['Normal']


def dialogue(doc,header,paragraph):
    head_f = doc.add_paragraph(header)
    head_f.paragraph_format.left_indent = Inches(2.0)
    body_f = doc.add_paragraph(paragraph.strip())
    body_f.paragraph_format.left_indent = Inches(1.0)
    head_f.style = doc.styles['Normal']
    body_f.style = doc.styles['Normal']


def convert(path):
    read, path = open_read(path)
    write = open_write()
    write.add_paragraph('FADE IN:')
    for para_obj in read.paragraphs:
        paragraph = [x for x in re.split('<|>',para_obj.text) if x]
        if len(paragraph) == 1:
            description(write,paragraph[0])
        elif len(paragraph) == 2:
            header = paragraph[0].strip().upper()
            if is_subheader(header):
                heading(write,header,paragraph[1])
            else:
                dialogue(write,header,paragraph[1])
    directory = path.split('/')
    if len(directory) > 1:
    #write.save('formated'+path)
        write.save('/'.join(directory[:-1])+'/format_'+directory[-1])
    else:
        write.save('format_'+path)


def driver():
    path = ''
    while path != 'exit':
        prompt = 'Enter document to convert, "exit" to end, and ' \
                    '"script.docx" is default.\n'
        path = input(prompt)
        path = path.strip().lower()
        if path == 'exit':
            return print('Program ended')
        convert(path)


driver()

# A python script that reads and formats scripts
