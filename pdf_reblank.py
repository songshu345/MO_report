import os

from PyPDF2 import PdfFileReader, PdfFileWriter

filework = 'D:/treasture/cloud_ama/report_automation/mo_pdf/'
outfilework = 'D:/treasture/cloud_ama/report_automation/'
file_folder = 'mo_pdf_reblank' # 去除空白页的pdf文件夹
os.makedirs(file_folder)

for pdf_file in os.listdir(filework):
    pdf_path = filework + pdf_file  # excel_path
    pdfReader = PdfFileReader(open(pdf_path, 'rb'))
    pdfFileWriter = PdfFileWriter()
    numPages = pdfReader.getNumPages()
    pagelist = (6, 7, 8, 9, 10, 11)  # 注意第一页的index为0.
    for index in range(0, numPages):
        if index not in pagelist:
            pageObj = pdfReader.getPage(index)
            pdfFileWriter.addPage(pageObj)
    pdf_file_str = pdf_file.split('.')[0]
    outfile = outfilework + file_folder + '/' + pdf_file_str + '_1' + '.pdf'
    pdfFileWriter.write(open(outfile, 'wb'))
