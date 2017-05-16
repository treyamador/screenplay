# a pythonic screenplay formatter
from docx import Document
from docx.shared import Inches
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
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
    style_margins(doc,0,0,1.5,1)
    return doc


def split_bracketed(paragraph,left,right):
    return [x for x in re.split(left+'|'+right,paragraph) if x]


def description(doc,paragraph):
    style_paragraph(doc,paragraph,0.0,0.0,None)


def is_subheader(header):
    return header == 'INT' or header == 'EXT' or header == 'SUB'


def heading(doc,header,body):
    heading = header+'. '+body.strip().upper()
    style_paragraph(doc,heading,0.0,0.0,None)


def dialogue(doc,header,paragraph):
    style_paragraph(doc,header,2.2,2.0,0)
    if paragraph.startswith('('):
        delim = paragraph.strip('(').split(')',1)
        style_paragraph(doc,'('+delim[0]+')',1.6,2.0,0)
        paragraph = delim[1].strip()
    style_paragraph(doc,paragraph,1.0,1.5,None)


def transition(doc,text):
    if not text.endswith(':'):
        text += ':'
    fmt = style_paragraph(doc,text,0.0,0.0,None)
    fmt.alignment = WD_ALIGN_PARAGRAPH.RIGHT


def page_breaks():
    pass


def style_margins(doc,top,bottom,left,right):
    for section in doc.sections:
        section.left_margin = Inches(left)
        section.right_margin = Inches(right)


def style_paragraph(doc,text,left,right,carriage):
    fmt = doc.add_paragraph(text)
    fmt.paragraph_format.left_indent = Inches(left)
    fmt.paragraph_format.right_indent = Inches(right)
    fmt.style = doc.styles['Normal']
    if carriage is not None:
        fmt.paragraph_format.space_after = Pt(carriage)
    return fmt


def convert(path):
    read, path = open_read(path)
    write = open_write()
    for para_obj in read.paragraphs:
        paragraph = split_bracketed(para_obj.text,'<','>')
        if len(paragraph) == 1:
            description(write,paragraph[0])
        elif len(paragraph) >= 2:
            header = paragraph[0].strip().upper()
            if is_subheader(header):
                heading(write,header,paragraph[1])
            elif header == 'TRAN':
                transition(write,paragraph[1].strip().upper())
            else:
                dialogue(write,header,paragraph[1].strip())
    directory = path.split('/')
    if len(directory) > 1:
        write.save('/'.join(directory[:-1])+'/format_'+directory[-1])
    else:
        write.save('format_'+path)


def help_prompt():
    help_msg = "\nThe current markup is '<' and '>'.\n" \
        "These can be changed by entering '--markup' and " \
        "the symbols separated by spaces.\n" \
        "Enter 'exit' to quit.\n"
    print(help_msg)


def driver():
    path = ''
    while path != 'exit':
        prompt = "\nEnter filepath of .docx to format or " \
                "--help' for instructions.\n"
        path = input(prompt).strip().lower()
        #path = path.strip().lower()
        if path == 'exit':
            return print('Program ended')
        elif path == '--help':
            help_prompt()
        else:
            convert(path)


driver()


# A pythonic script that reads and formats scripts!
