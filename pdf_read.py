
import os
import re
import pdfplumber
import openpyxl
from tqdm import tqdm
import shutil  # 添加shutil库

def extract_info_from_pdf(pdf_path):

    with pdfplumber.open(pdf_path) as pdf:
        first_page = pdf.pages[0]
        text = first_page.extract_text()
    #print(text)

    # 匹配题目
    #title_pattern = re.compile(r'题[目|名][:：]\s*(.*?)\s*(?=学号[:：]|姓名[:：]|$)', re.DOTALL)
    title_pattern = re.compile(r'题目[:：](.*?)学\s*号[:：]', re.DOTALL)
    title = re.findall(title_pattern, text)
    if title:
        title = title[0].replace("\n", "").strip()
    else:
        title = "未提取到题目信息"
        #raise ValueError("未提取到题目信息")
    #print(re.findall(title_pattern, text))

    # 匹配学号
    id_pattern = re.compile(r'学\s*号[:：](.*?)\n')
    student_id = re.findall(id_pattern, text)
    if student_id:
        student_id = student_id[0].replace(" ", "").strip()
    else:
        student_id = "未提取到学号信息"
       #raise ValueError("未提取到学号信息")

    # 匹配姓名
    name_pattern = re.compile(r'姓\s*[名|氏][:：](.*?)\n')
    name = re.findall(name_pattern, text)
    if name:
        name = name[0].replace(" ", "").strip()
    else:
        name = "未提取到姓名信息"
        #raise ValueError("未提取到姓名信息")

    # 匹配专业
    major_pattern = re.compile(r'专\s*业[:：](.*?)\n')
    major = re.findall(major_pattern, text)
    if major:
        major = major[0].replace(" ", "").strip()
    else:
        major_pattern = re.compile(r'专业领域[:：](.*?)\n')
        major = re.findall(major_pattern, text)
        if major:
            major = major[0].replace(" ", "").strip()
        else:
            major = "未提取到专业信息"
            #raise ValueError("未提取到专业信息")

    # 提取导师
    tutor_pattern = re.compile(r'[导指]\s*[师教][:：](.*?)\n')
    tutor = re.findall(tutor_pattern, text)
    if tutor:
        tutor = tutor[0].replace(" ", "").strip()
    else:
        tutor = "未提取到导师信息"
        #raise ValueError("未提取到导师信息")

    # 匹配学院
    school_pattern = re.compile(r'学\s*[院|校][:：](.*?)\n')
    school = re.findall(school_pattern, text)
    if school:
        school = school[0].replace(" ", "").strip()
    else:
        school_pattern = re.compile(r'学\s*[院|校][:：](.*?)学院', re.DOTALL)
        school = re.findall(school_pattern, text)
        if school:
            school = school[0].replace(" ", "").strip() + '学院'
        else:
            school = "未提取到学院信息"
        #raise ValueError("未提取到学院信息")

    # 提取日期
    date_pattern = re.compile(r'\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日')
    date = re.findall(date_pattern, text)
    if date:
        date = date[0].replace(" ", "").strip()
    else:
        date_pattern = re.compile(r'\d{4}\s*年\s*\d{1,2}\s*月')
        date = re.findall(date_pattern, text)
        if date:
            date = date[0].replace(" ", "").strip()
        else:
            date = "未提取到日期信息"
            #raise ValueError("未提取到日期信息")

    return title, student_id, name, major, tutor, school, date




def process_pdfs_in_folder(folder_path):
    output_path = os.path.join(folder_path, "output.xlsx")
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(["题目", "学号", "姓名", "专业", "导师", "学院", "日期"])

    success_count = 0
    failure_count = 0

    #for file in os.listdir(folder_path):
    # 使用tqdm添加进度条
    for file in tqdm(os.listdir(folder_path), desc="Processing PDFs"):
        if file.endswith(".pdf") or file.endswith(".PDF"):
            pdf_path = os.path.join(folder_path, file)
            try:
                info = extract_info_from_pdf(pdf_path)
                title, student_id, name, major, tutor, school, date = info

                if (title != "未提取到题目信息") and (student_id != "未提取到学号信息") and (name != "未提取到姓名信息") and (major != "未提取到专业信息") and (tutor != "未提取到导师信息") and (school != "未提取到学校信息") and (date != "未提取到日期信息"):
                    sheet.append([title, student_id, name, major, tutor, school, date])
                    new_name = os.path.join(folder_path, student_id + ".pdf")
                    if not os.path.exists(new_name):
                        os.rename(pdf_path, new_name)
                    shutil.move(new_name, os.path.join(success_pdf_folder, student_id + ".pdf"))  # 添加移动文件的代码
                    success_count += 1
                else:
                    print(f"警告：{file} 未能提取所有信息")
                    failure_count += 1
            except Exception as e:
                print(f"error:提取PDF文件 {pdf_path} 的信息时出现错误：{str(e)}")
                failure_count += 1
    workbook.save(output_path)
    print(f"\n共有{success_count + failure_count}个PDF文件，成功提取{success_count}个，提取失败{failure_count}个。")


pdf_folder = "E:\\图书馆数据\\pdf_extracion\\pdf"
success_pdf_folder = "E:\\图书馆数据\\pdf_extracion\\succcess"

process_pdfs_in_folder(pdf_folder)

'''
    使用相对路径"data"来引用data文件夹，并使用os.path.join函数来创建文件的完整路径。
    无论将my_project文件夹移动到哪个位置，代码中使用的相对路径部分都将指向相同的文件和文件夹。
'''

'''
print("题目: ", title)
print("学号: ", student_id)
print("姓名: ", name)
print("专业: ", major)
print("导师: ", tutor)
print("学院: ", school)
print("日期: ", date)
'''