from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from io import StringIO

def convert_pdf_to_txt(path, pages):
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    codec = 'utf-8'

    laparams=LAParams(all_texts=True
                      , detect_vertical=True)#, 
##                      line_overlap=0.5, char_margin=1000.0, #set char_margin to a large number
##                      line_margin=0.5, word_margin=2,
##                      boxes_flow=1)
    device = TextConverter(rsrcmgr, retstr, laparams=laparams)
    fp = open(path, 'rb')
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    password = ""
    maxpages = 0
    caching = True
    pagenos=set(pages)

    for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages, password=password,caching=caching, check_extractable=True):
        interpreter.process_page(page)

    text = retstr.getvalue()

    fp.close()
    device.close()
    retstr.close()
    return text

pdf_text_page6 = convert_pdf_to_txt("C:/Users/Ecava/Downloads/Schedule 10.8.23-10.14.23.pdf",
                                    pages=[0])
if "CAVANNA" in pdf_text_page6:
    print("stage 1")
#print(pdf_text_page6)


