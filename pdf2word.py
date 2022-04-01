# -*- coding = utf-8 -*-
# @Time : 2021/9/17 10:56
# @Author:Chenguang Pan
'''
This py file mainly focuses on convert the pdf to docx and keep
the original format simultaneously.

Please install the pdf2docx package first, see this url for more details:https://github.com/dothinking/pdf2docx

'''


import os
from pdf2docx import Converter
import time


# set the default path
os.chdir('/Users/panpeter/Desktop/pythonProject/word2pdf/OriginalPDF')
# store the default path
file_dir = os.getcwd()

print('-' * 3 + 'stage 1:pdf读取中......' + '-' * 3)

# adding a timing module
time_start = time.time()

# iterate all files
for root, dirs, files in os.walk(file_dir):
    pdf_files = files

pdf_files.remove('.DS_Store')
print(pdf_files)

print('-' * 3 + 'stage 2:pdf读取完毕' + '-' * 3)
print('-' * 3 + 'stage 3:pdf转换开始...' + '-' * 3)


'''
# to remove the useless file
for pdf in pdf_files:
    if pdf[-3] != 'pdf':
        pdf_files.remove(pdf)

print(pdf_files)
'''


# convert pdf to docx
for pdf in pdf_files:
    doc_file = pdf.replace('.pdf', '.docx')
    cvt = Converter(pdf)
    cvt.convert(doc_file)
    cvt.close()

print('-' * 3 + 'stage 4:pdf转换完毕...' + '-' * 3)
time_end = time.time()

time_all = time_end - time_start
time_around = round(time_all, 2)
print("转换总用时：" + str(time_around) + "秒")
