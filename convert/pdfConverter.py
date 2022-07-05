import os
from docx2pdf import convert
from PyPDF2 import PdfFileMerger

mainFolderPath = os.getcwd()
filesToConvertPath = os.getcwd() + r'\docxFiles'
folderWithPdfFIles = os.getcwd() + r'\pdfFiles'

# Potential com_error, script sets the path to system32 folder, Documents, etc
os.chdir(filesToConvertPath)

for i in range(len(os.listdir(filesToConvertPath))):
    convert(os.listdir(filesToConvertPath)[i], folderWithPdfFIles)

os.chdir(folderWithPdfFIles)
merger = PdfFileMerger()

for root, dirs, fileNames in os.walk(folderWithPdfFIles):
    for fileName in fileNames:
        merger.append(folderWithPdfFIles + r'\\' + fileName)

os.chdir(mainFolderPath)
merger.write('combinedFiles.pdf')
merger.close()