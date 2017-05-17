# a pythonic screenplay formatter
from docx import Document
from docx.shared import Inches
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re, os


def open_read(path):
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


# accept transform here
def description(doc,paragraph,keys):
    paragraph = transform(keys,paragraph)
    style_paragraph(doc,paragraph,0.0,0.0,None)


def is_subheader(header):
    return header == 'INT' or header == 'EXT' or header == 'SUB'


# accept transform here
def heading(doc,header,text,keys):
    desc = transform(keys,header)
    heading = header+'. '+desc.strip().upper()
    style_paragraph(doc,heading,0.0,0.0,None)


# accept transform here
def dialogue(doc,header,paragraph,tags,keys):
    header = transform(tags,header)
    paragraph = transform(keys,paragraph)
    style_paragraph(doc,header,2.2,2.0,0)
    if paragraph.startswith('('):
        delim = paragraph.strip('(').split(')',1)
        style_paragraph(doc,'('+delim[0]+')',1.6,2.0,0)
        paragraph = delim[1].strip()
    style_paragraph(doc,paragraph,1.0,1.5,None)


# accept in transform here
def transition(doc,text,keys):
    text = transform(keys,text)
    if not text.endswith(':'):
        text += ':'
    fmt = style_paragraph(doc,text,0.0,0.0,None)
    fmt.alignment = WD_ALIGN_PARAGRAPH.RIGHT


def add_tags(trans,text):
    key,value = [x.strip().upper() for x in text.split('=')][:2]
    trans[key] = value


def add_keys(trans,text):
    key,value = [x.strip() for x in text.split('=')][:2]
    trans[key] = value


#def replace_tag(tags,header):
#    pass
    #for key,value in tags.items():
    #    if key in header:
    #        header = re.sub(key,)
    #if header in tags.keys():
    #    header = tags[header]
    #return header


# keep an eye on this, make sure it continues to look good
def transform(trans,text):
    for key, value in trans.items():
        text = re.sub('((?<= )|^)'+key+'((?=[ .?!])|$)',value,text)
    return text


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


def save_doc(doc,path):
    name = path.split('.')
    doc.save('.'.join(name[:-1])+'_formatted.'+name[-1])


def convert(path):
    read, path = open_read(path)
    write = open_write()
    tags,keys = {},{}
    for para_obj in read.paragraphs:
        paragraph = split_bracketed(para_obj.text,'<','>')
        header = paragraph[0].strip().upper()
        if len(paragraph) == 1:
            description(write,paragraph[0],keys)
        elif is_subheader(header):
            heading(write,header,paragraph[1],keys)
        elif header == 'TRAN':
            transition(write,paragraph[1].strip().upper(),keys)
        elif header == 'TAG':
            add_tags(tags,paragraph[1])
        elif header == 'KEY':
            add_keys(keys,paragraph[1])
        else:
            dialogue(write,header,paragraph[1].strip(),tags,keys)
    save_doc(write,path)


def validity(path):
    if path == '':
        path = 'script.docx'
    elif not path.endswith('.docx'):
        path += '.docx'
    if path not in os.listdir():
        return print('That file does not exist.')
    return path


def help_prompt():
    help_msg = "\nThe current markup is '<' and '>'.\n" \
        "These can be changed by entering '--markup' and " \
        "the symbols separated by spaces.\n" \
        "Enter 'exit' to quit.\n"
    print(help_msg)


def driver():
    while True:
        prompt = "\nEnter filepath of .docx to format or " \
                "--help' for instructions.\n"
        path = input(prompt).strip()
        if path == 'exit':
            return print('Program ended')
        elif path == '--help':
            help_prompt()
        else:
            path = validity(path)
            if path:
                convert(path)


driver()


# A pythonic script that reads and formats scripts!
