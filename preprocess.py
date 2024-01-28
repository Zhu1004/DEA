import mysql.connector
import RedsealRomove.main
import glob
import os
import logging
from pdf2docx import Converter
import extractimgfromword
import RedsealRomove.main
import newtestnyf
import ocr
import end
import openpyxl
from openpyxl import Workbook
import test_delete
from flask import Flask, request
import json
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
def pdf_to_word(pdf_file_path, word_file_path):
    cv = Converter(pdf_file_path)
    cv.convert(word_file_path)
    cv.close()
def dea(tag,pdf_file,round,relative_path):


    word_folder = "wordout"  # Word文件输出目录路径
    logging.getLogger().setLevel(logging.ERROR)
    tasks = []
    pdf_file_name = os.path.basename(pdf_file)
    file_name = os.path.splitext(pdf_file_name)[0]
    word_file = os.path.join(word_folder + "/example.docx")
    #print(word_file)
    #print("正在处理: ", pdf_file_name)
    pdf_to_word(pdf_file, word_file)
    tasks.append(True)

    while True:
        exit_flag = True
        for task in tasks:
            if not task:
                exit_flag = False
        if exit_flag:
            print("前置转换完成")

            temp=extractimgfromword.run_1()
            if temp==0:
                RedsealRomove.main.run_2()

                ocr.run_3()

                end.run_4()

            test_delete.rubbish_delete()
            newtestnyf.run_5(relative_path,round)
            break
        else:
            print("前置转换出现问题")



app = Flask(__name__)
@app.route('/preprocess', methods=['POST'])
def preprocess():
    root_path = "E:/pcpy_file/files/supbiaoshu_jiemi/"
    data = request.get_json()
    tag = data["tag"]
    filepath = data["filePath"]
    filepath = filepath.replace("\\", "/")
    if tag is None or filepath is None:
        returndatas = {
            "code": "000002",
            "message": "fail，tag is None or filepath is None",
            "data": None
        }
        return json.dumps(returndatas)
    try:
        db = mysql.connector.connect(
            host='127.0.0.1',
            port='3306',
            user='root',
            password='111111',
            database='dea'
        )
    except:
        returndatas = {
            "code": "000002",
            "message": "fail to connect the database",
            "data": None
        }
        return json.dumps(returndatas)

    #先，分割tag，清空并写入excel
    #遍历所有pdf，连续跑
    excel_path = "data_input.xlsx"
    if os.path.exists(excel_path):
        os.remove(excel_path)
    workbook = Workbook()
    workbook.save(excel_path)
    # 加入excel
    wb = openpyxl.load_workbook("data_input.xlsx")
    sheet = wb.active
    for idx, tags in enumerate(tag.split(",")):
        cursor = db.cursor()
        sql = "select * from deatest where tag=" + tags + ";"
        print(sql)
        cursor.execute(sql)
        data = cursor.fetchone()
        if data is not None:  # 如果为空，切掉这个tag,
            print(tags)
            cell = "B{}".format(idx + 1)
            sheet[cell] = tags
    wb.save("data_input.xlsx")
    file_names = get_file_names(root_path+filepath)
    for pdf_file in file_names:
        pdf_file = pdf_file.replace("\\", "/")
        print(pdf_file)
        relative_path = os.path.relpath(pdf_file, root_path)
        relative_path=relative_path.replace("\\", "/")
        print(relative_path)
        dea(None, pdf_file, 2, relative_path)


    returndatas = {
        "code": "000004",
        "message": "preprocess",
        "data": None
    }
    return json.dumps(returndatas)



if __name__ == "__main__":
    app.run(port=5000, debug=True)