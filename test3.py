from pdfminer.layout import LAParams, LTTextBox
import pandas as pd
import re
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LTTextContainer

#fixed params and initilization
empName = "CAVANNA"
search = 'CAVANNA, Eric\n'

#create a dataframe
columns = ["row", "column", "value"]

#df = pd.DataFrame(columns=columns)

fp = open("C:/Users/Ecava/Downloads/Schedule.pdf", 'rb')
rsrcmgr = PDFResourceManager()
laparams = LAParams()
device = PDFPageAggregator(rsrcmgr, laparams=laparams)
interpreter = PDFPageInterpreter(rsrcmgr, device)
pages = PDFPage.get_pages(fp)

for page in pages:
    print('Processing next page...')
    interpreter.process_page(page)
    layout = device.get_result()
    df = []
    for lobj in layout:
        if isinstance(lobj, LTTextBox):
            x, y, text = lobj.bbox[0], lobj.bbox[3], lobj.get_text()
            #print('%r %s' % ((x, y), text))
            rowData = {"column": x, "row": y, "value": text}
            df.append(rowData)

    pgdf = pd.DataFrame(df)
    name_match = pgdf[pgdf['value']==search]
    start_row = name_match.row.values[0]
    start_col = name_match.column.values[0]
    print(start_col)
    print(start_row)
    pgdf.value = pgdf.value.str.strip()
    print(pgdf[(pgdf.column > 40)&(pgdf['value']!= "")].head(60))
    
    

    
    

    
    
    
