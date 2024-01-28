import os


def get_file_names(folder_path):
    """遍历文件夹及子文件夹，读取所有文件的文件名"""
    file_names = []
    # 遍历文件夹及子文件夹
    for root, dirs, files in os.walk(folder_path):
        # 遍历当前文件夹下的所有文件
        for file_name in files:
            # 将文件名添加到列表中
            # print(root+"\\"+file_name)
            file_names.append(root+"\\"+file_name)

    return file_names


# 示例用法
folder_path = "E:/pcpy_file/files/supbiaoshu_jiemi/23CS02"  # 文件夹路径

file_names = get_file_names(folder_path)
for file_name in file_names:
    print(file_name)
