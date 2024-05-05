import os
from win32com.client import Dispatch

path = os.getcwd()

old_file_path = os.path.abspath('D:/14482/PyCharmProject/.A-存储的项目/2月1号的读书报告会.pptx')
new_file_path = os.path.abspath('D:/14482/PyCharmProject/.A-存储的项目/2月1号的读书报告会.pdf')

powerpoint = Dispatch('PowerPoint.Application')
presentation = powerpoint.Presentations.Open(old_file_path)

wdFormatPDF = 32
presentation.SaveAs(new_file_path, FileFormat=wdFormatPDF)

presentation.Close()
powerpoint.Quit()

print('successfully')