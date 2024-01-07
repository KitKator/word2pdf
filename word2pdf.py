import os
import shutil

import win32com.client
from docx import Document
from tqdm import tqdm

def convert_to_pdf(doc_file, pdf_file):
    #try:
        # 打开 Word 文档
        word = win32com.client.Dispatch('Word.Application')
        doc = word.Documents.Open(doc_file)
        doc.SaveAs(pdf_file, FileFormat=17)
        doc.Close()
        word.Quit()
    #except Exception as e:
        #print(f'Error converting {doc_file} to {pdf_file}: {str(e)}')
'''
首先使用 win32com 模块创建了一个 Word.Application 对象，
然后使用 Open 方法打开了要转换的 Word 文档。
接着，调用 SaveAs 方法将文档保存为 PDF 文件，FileFormat=17 表示将文档保存为 PDF 格式。
最后，调用 Close 方法关闭 Word 文档，以及调用 Quit 方法退出 Word 应用程序。
'''

# 设置相对路径的起点为当前脚本所在目录
word_folder = "E:\\图书馆数据\\word"
pdf_folder = "E:\\图书馆数据\\word_pdf"
input_folder_path = word_folder
word_good_path = "E:\\图书馆数据\\word_pdf\\word_good"
word_bad_path = "E:\\图书馆数据\\word_pdf\\word_bab"
pdf_path = "E:\\图书馆数据\\word_pdf\\pdf"

# 获取要转换的文件名列表
file_names = [f for f in os.listdir(input_folder_path) if f.endswith('.docx') or f.endswith('.doc') or f.endswith('.DOC')]
success_count = 0
error_count = 0

# 在循环中加入进度条
for file_name in tqdm(file_names, desc='Converting files', unit='file'):
    # 构造 Word 文档和 PDF 文件的绝对路径
    doc_file = os.path.join(input_folder_path, file_name)
    pdf_file = os.path.join(pdf_path, os.path.splitext(file_name)[0] + '.pdf')

    # 如果 PDF 文件已经存在，则跳过该文件
    if os.path.exists(pdf_file):
        print(f'Skipping {doc_file}: PDF file already exists')
        continue

    try:
        # 执行转换操作
        convert_to_pdf(doc_file, pdf_file)
        success_count += 1
        shutil.move(doc_file, os.path.join(word_good_path, file_name))
    except Exception as e:
        print(f'Error converting {doc_file} to {pdf_file}: {str(e)}')
        error_count += 1
        #shutil.move(doc_file, os.path.join(word_bad_path, file_name))

total_count = len(file_names)
print(f'total_word: {total_count}, Success: {success_count}, Error: {error_count}')


'''
使用 tqdm 库中的 tqdm 函数将要遍历的文件名列表封装成一个进度条对象。
在循环遍历文件名时，将文件名作为进度条的单位，并使用 desc 参数指定进度条的描述信息。
'''
'''
如果打开 Word 文档失败，或者保存 PDF 文件时发生错误，程序将输出错误信息并跳过当前文件。
如果 PDF 文件已经存在，则程序会输出一条警告信息并跳过该文件。
'''