import PyPDF2
def extract_text_pdf(pdf):
    # creating a pdf file object
    pdfFileObj = open(pdf, 'rb')
    
    # creating a pdf reader object
    pdfReader = PyPDF2.PdfReader(pdfFileObj)
    
    # printing number of pages in pdf file
    # print(len(pdfReader.pages))
    i = 0
    num_pages = len(pdfReader.pages)
    pdf_text = ''

    while i < num_pages:
        # creating a page object
        pageObj = pdfReader.pages[i]
    
        # extracting text from page
        if i == 0:
            pdf_text = pdf_text + '' + pageObj.extract_text()
        else: 
            pdf_text = pdf_text + '\n\n' + pageObj.extract_text()
        i = i + 1
    # closing the pdf file object
    pdfFileObj.close()

    return pdf_text