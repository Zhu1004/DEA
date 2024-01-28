import json
import mysql.connector
from flask import Flask, request
app = Flask(__name__)
@app.route('/front', methods=['POST'])
def register():
    data = request.get_json()
    tag = data["tag"]
    filepath = data["filePath"]
    filepath=filepath.replace("\\", "/")
    print(filepath)
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
    #检查tag是否在模板库中
    cursor = db.cursor()
    sql = "select * from deatest where tag=" + tag + ";"
    print(sql)
    cursor.execute(sql)
    data = cursor.fetchone()
    if data is None:
        returndatas = {
            "code": "000002",
            "message": "the tag is not existing",
            "data": None
        }
        return json.dumps(returndatas)

    cursor = db.cursor()
    sql = "select page from convert_result where tag=" + tag + " and pdfpath='" + filepath + "';"
    print(sql)
    cursor.execute(sql)
    data = cursor.fetchone()

    if data:
        returndata={
            "tag": tag,
            "page": data
        }
        returndatas = {
            "code": "000000",
            "message": "success",
            "data": returndata
        }
        return json.dumps(returndatas)
    else:
        try:
            cursor.execute("insert into convert_condition values(" + tag + ",'" + filepath + "',1);")
            db.commit()
            print("insert into convert_condition values(" + tag + ",'" + filepath + "',1);")
        except:
            print("请勿重复点击")
        returndatas = {
            "code": "000001",
            "message": "converting",
            "data": None
        }
        return json.dumps(returndatas)



if __name__ == "__main__":

    app.run(port=5000, debug=True)
