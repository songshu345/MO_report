from win32com.client import DispatchEx
import os
from PyPDF2 import PdfFileReader, PdfFileWriter

# 文件夹路径
excel_work = 'D:/treasture/cloud_ama/report_automation/mo_masking/'  # 脱敏后的Excel文件所在文件夹
outfilework = 'D:/treasture/cloud_ama/report_automation/'
file_folder = 'mo_pdf_reblank'  # 去除空白页的pdf文件夹
pdf_folder = 'mo_pdf'  # Excel转为的pdf文件将要存储的文件夹

os.makedirs(pdf_folder)
os.makedirs(file_folder)

# Excel转为Pdf
for excel_file in os.listdir(excel_work):
    excel_path = excel_work + excel_file  # excel_path
    pdf_file = excel_file.split('.')[0]
    pdf_work = outfilework + pdf_folder + '/'
    pdf_path = pdf_work + pdf_file  # pdf_path
    xlApp = DispatchEx("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = 0
    books = xlApp.Workbooks.Open(excel_path, False)
    books.ExportAsFixedFormat(0, pdf_path)
    books.Close(False)
    xlApp.Quit()
#  去除pdf中的空白页（空白页已知）
for pdf_file in os.listdir(pdf_work):
    pdf_path = pdf_work + pdf_file  # excel_path
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
