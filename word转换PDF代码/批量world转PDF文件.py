import os
from win32com.client import Dispatch

path = os.getcwd()

old_file_path = os.path.abspath('D:/14482/pythonProject/.A储存项目/草稿1.docx')
new_file_path = os.path.abspath('D:/14482/pythonProject/.A储存项目/草稿1.pdf')

word = Dispatch('Word.Application')
doc = word.Documents.Open(old_file_path)
wdFormatPDF = 17
doc.SaveAs(new_file_path, FileFormat=wdFormatPDF)
doc.Close()
word.Quit()

print('successfully')