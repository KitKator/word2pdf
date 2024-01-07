import os

folder_path = 'E:\\图书馆数据\\pdf_extracion\\success'  # 替换为实际文件夹路径
pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]


# 从文本文件中读取学号数据
with open('E:\\图书馆数据\\pdf_extracion\\numbers.txt', 'r', encoding='utf-8') as file:
    student_numbers = set(file.read().splitlines())

# 将文件夹中的PDF文件名提取学号
pdf_numbers = set([pdf_file.split('.')[0] for pdf_file in pdf_files])

# 查找缺失的学号
missing_numbers = student_numbers - pdf_numbers
miss_2 = pdf_numbers - student_numbers
print("文件中有但是")
print(miss_2)
if missing_numbers:
    print("文件夹中缺失的学号：")
    for number in missing_numbers:
        print(number)

