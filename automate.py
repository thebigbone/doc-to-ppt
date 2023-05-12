from docx2pdf import convert
import os
from PyPDF2 import PdfReader
from pptx import Presentation

try:
    docx_file = input("Enter a docx file: ")
    convert(docx_file, "output.pdf")
    pdf_file = 'output.pdf'
    print('converted to pdf.')
except:
    print("file not found. please enter the full path or put the file in the same dir")
    

def pdf_to_ppt(pdf_path, ppt_path):
    try:
        pdf = PdfReader(open(pdf_path, 'rb'))

        prs = Presentation()
        for page_num in range(len(pdf.pages)):
            page = pdf.pages[page_num]
            text = page.extract_text()

            slide_layout = prs.slide_layouts[1]
            slide = prs.slides.add_slide(slide_layout)

            text_box = slide.shapes.add_textbox(0, 0, prs.slide_width, prs.slide_height)
            text_frame = text_box.text_frame
            text_frame.text = text
        prs.save(ppt_path)
    except:
        print('something went wrong!')

pdf_file = 'output.pdf'
ppt_file = 'output.pptx'

pdf_to_ppt(pdf_file, ppt_file)
print('conversion successful. file name: ' + ppt_file)