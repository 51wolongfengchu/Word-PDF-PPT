import os
from win32com.client import Dispatch

# 设置要转换的PPT文件所在的目录
input_directory = 'C:/Users/14482/Desktop/未整理'
# 设置转换后的PDF文件保存的目录
output_directory = 'C:/Users/14482/Desktop/整理完成/converted'

# 确保输出目录存在
if not os.path.exists(output_directory):
    os.makedirs(output_directory)

# 列出目录中的所有文件
for filename in os.listdir(input_directory):
    # 检查文件扩展名是否为.pptx
    if filename.endswith('.pptx'):
        # 构建完整的文件路径
        old_file_path = os.path.abspath(os.path.join(input_directory, filename))
        # 构建新的PDF文件路径
        new_file_path = os.path.abspath(os.path.join(output_directory, f"{os.path.splitext(filename)[0]}.pdf"))

        # 打印转换信息
        print(f"Converting {filename} to PDF...")

        # 调用PowerPoint进行转换
        powerpoint = Dispatch('PowerPoint.Application')
        presentation = powerpoint.Presentations.Open(old_file_path)

        # 保存为PDF
        presentation.SaveAs(new_file_path, FileFormat=32)

        # 关闭PowerPoint文件
        presentation.Close()

        # 打印转换成功信息
        print(f"Successfully converted {filename} to PDF.")

# 关闭PowerPoint应用程序
powerpoint.Quit()

print('successfully')
