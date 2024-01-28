import mysql.connector
import time
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
from datetime import datetime
import test_delete
def pdf_to_word(pdf_file_path, word_file_path):
    cv = Converter(pdf_file_path)
    cv.convert(word_file_path)
    cv.close()
def dea(tag,pdf_file,round,relative_path):
    excel_path="data_input.xlsx"
    if os.path.exists(excel_path):
        os.remove(excel_path)
    workbook = Workbook()
    workbook.save(excel_path)
    # 加入excel
    wb = openpyxl.load_workbook("data_input.xlsx")
    sheet = wb.active
    col_tagin = sheet['B']  # 输入的细则tag
    col_tagin[0].value = tag
    wb.save("data_input.xlsx")

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

def after():
    root_path = "E:/pcpy_file/files/supbiaoshu_jiemi/"
    while True:
        try:
            db = mysql.connector.connect(
                host='127.0.0.1',
                port='3306',
                user='root',
                password='111111',
                database='dea'
            )
        except:
            print("连接数据库失败！")
        cursor = db.cursor()
        sql = "select tag,pdfpath from convert_condition where state=1"
        print("检查任务表中")
        cursor.execute(sql)
        data = cursor.fetchone()
        if data is not None:
            now = datetime.now()

            # 格式化输出日期和时间
            current_time = now.strftime("%Y-%m-%d %H:%M:%S")

            print("[", current_time, "]捕获到新的处理任务")
            print()
            print(data[0],root_path+data[1])
            complete_path=root_path+data[1]
            dea(data[0],complete_path,1,data[1])
        else:
            now = datetime.now()

            # 格式化输出日期和时间
            current_time = now.strftime("%Y-%m-%d %H:%M:%S")

            print("[", current_time,"]没有新的处理任务")
            print()
            time.sleep(20)

after()