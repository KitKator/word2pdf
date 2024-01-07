import os
import shutil
import pdfplumber
from tqdm import tqdm


current_folder = os.path.dirname(os.path.abspath(__file__))
data_folder = "data"
input_folder = os.path.join(current_folder, data_folder, "all_pdf")
output_folder = os.path.join(current_folder, data_folder, "NOT_open_pdf")


total_files = 0
openable_files = 0
unopenable_files = 0

pdf_files = [f for f in os.listdir(input_folder) if f.endswith('.pdf')]

for file in tqdm(pdf_files, desc='Checking PDFs', unit='file'):
    total_files += 1

    try:
        with pdfplumber.open(os.path.join(input_folder, file)) as pdf:
            pass
        openable_files += 1

    except pdfplumber.pdf.PDFSyntaxError:
        unopenable_files += 1
        shutil.move(os.path.join(input_folder, file), os.path.join(output_folder, file))

print(f'Total files: {total_files}')
print(f'Openable files: {openable_files}')
print(f'Unopenable files: {unopenable_files}')
