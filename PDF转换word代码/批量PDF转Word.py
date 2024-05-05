import os
from pdf2docx import Converter

def pdf_to_word(pdf_file_path, word_file_path):
    # 初始化转换器
    cv = Converter(pdf_file_path)

    # 转换PDF文件
    cv.convert(word_file_path, start=0, end=None)

    # 输出成功信息
    print(f"PDF文件 '{pdf_file_path}' 已成功转换为Word文档 '{word_file_path}'")

def batch_convert(pdf_folder, word_folder):
    # 确保输出文件夹存在
    if not os.path.exists(word_folder):
        os.makedirs(word_folder)

    # 遍历PDF文件夹中的所有PDF文件
    for filename in os.listdir(pdf_folder):
        if filename.endswith('.pdf'):
            pdf_file_path = os.path.join(pdf_folder, filename)
            word_file_path = os.path.join(word_folder, f"{os.path.splitext(filename)[0]}.docx")

            # 执行转换，并捕获可能出现的异常
            try:
                pdf_to_word(pdf_file_path, word_file_path)
            except Exception as e:
                print(f"转换PDF文件时出错: {e}")

# 使用示例
pdf_folder = 'C:/Users/14482/Desktop/Organoids-literature'  # 这里替换为您的PDF文件所在的文件夹路径
word_folder = 'C:/Users/14482/Desktop/Organoids-literature/converted'  # 这里替换为您希望生成的Word文档所在的文件夹路径
batch_convert(pdf_folder, word_folder)
