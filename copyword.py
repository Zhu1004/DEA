import os
import shutil

def copy_file(source_file, destination_path, new_filename):
    # 获取源文件的文件名
    filename = os.path.basename(source_file)
    # 构建目标文件的完整路径
    destination_file = os.path.join(destination_path, new_filename)

    shutil.copy2(source_file, destination_file)

# 示例用法
source_file = 'path/to/source/file.txt'  # 源文件路径
destination_path = 'path/to/destination/'  # 目标路径，不包括文件名
new_filename = 'new_file.txt'  # 复制后的新文件名

copy_file(source_file, destination_path, new_filename)