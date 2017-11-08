#coding:utf-8
import xlwt
from flask import Flask, request, jsonify, render_template,send_file,make_response
from flask_script import Manager 
import sqlite3
import json
import os

def create_app():
  app = Flask(__name__)
  app.config['JSON_AS_ASCII'] = False
  return app

application = create_app()

def tojson(listdata):                                                                            # 返回json格式数据
    senddata = {}
    senddata = listdata
    return  jsonify(senddata)

@application.route('/student')
def index():
    return render_template('student.html')

@application.route('/export')                                                                    # 将数据库信息转换为excel
def export():                                 
    try:             
        os.remove(os.path.join(app.static_folder, 'ClassBook.xls'))                              
        conn = sqlite3.connect('test.db')                                                    # 连接数据库
        cursor = conn.cursor()
        count = cursor.execute("select* from student")                                      # 执行sql语句

        results = cursor.fetchall()                                                          # 搜取所有结果
        table_name = 'student'
        fields = cursor.description                                                          # 获取sqlite3里面的数据字段名称
        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet('table_'+table_name,cell_overwrite_ok=True)

        for field in range(0,len(fields)):                                                   # 写上字段信息
            sheet.write(0,field,fields[field][0])

        row = 1 
        col = 0
        for row in range(1,len(results)+1):                                                  # 获取并写入数据段信息
            for col in range(0,len(fields)):
                sheet.write(row,col,'%s'%results[row-1][col])

        workbook.save('ClassBook.xls')
        response = make_response(send_file("ClassBook.xls"))
        response.headers["Content-Disposition"] = "attachment; filename=ClassBook.xls;"
        cursor.close()                                                                       # 关闭数据库连接
        conn.commit()
        conn.close()
        return response
    except: 
        return jsonify({'finished': 'false'})

@application.route('/show', methods = ['POST','GET'])                                            # 显示所有学生信息
def show_stu():
    conn = sqlite3.connect('test.db')
    cursor = conn.cursor()
    try:
        cursor.execute("select* from student")
        listdata = []
        results = cursor.fetchall()
        for row in results:
            one_record = {}
            one_record['name'] = row[0]
            one_record['address'] = row[1]
            one_record['phone'] = row[2]
            one_record['wechat'] = row[3]
            one_record['mail'] = row[4]
            one_record['qq'] = row[5]
            one_record['note'] = row[6]
            listdata.append(one_record)    
        
        cursor.close()
        conn.commit()
        conn.close()
        return tojson(listdata)
    except:
        print("Error: unable to fecth data")
        return jsonify({'finished': 'false'})

@application.route('/add', methods = ['POST'])                                             # 添加学生信息
def add_stu():
    conn = sqlite3.connect('test.db')
    cursor = conn.cursor()
    data = request.get_data()
    data = data.decode("utf-8")
    dict = json.loads(data)
    values = (dict['name'],dict['address'],dict['phone'],dict['wechat'],
            dict['mail'],dict['qq'],dict['note'])                                                  # 接受前端发送的信息
    print(values)
    try:
        cursor.execute('''insert into student(name,address,phone,wechat,mail,qq,note) 
                       values(?,?,?,?,?,?,?)''',values)
        cursor.close()
        conn.commit()
        conn.close()
        return jsonify({'finished': 'true'})
    except:
        print("Error: unable to insert data")
        return jsonify({'finished': 'false'})

@application.route('/delete', methods = ['POST'])                                          # 删除学生信息                 
def delete_stu():
    conn = sqlite3.connect('test.db')
    cursor = conn.cursor()
    data = request.get_data()
    data = data.decode("utf-8")
    dict = json.loads(data) 
    phone = (dict['phone'],)
    try:
        cursor.execute("DELETE FROM student WHERE phone = ?",phone)
        cursor.close()
        conn.commit()
        conn.close()
        return jsonify({'finished': 'true'})
    except:
        print("Error: unable to delete data")
        return jsonify({'finished': 'false'})
      
@application.route('/change', methods = ['POST'])                                          # 更改学生信息
def change_stu():
    conn = sqlite3.connect('test.db')
    cursor = conn.cursor()
    conn = sqlite3.connect('test.db')
    cursor = conn.cursor()
    data = request.get_data()
    data = data.decode("utf-8")
    dict = json.loads(data)
    values = (dict['name'],dict['address'],dict['phone'],dict['wechat'],dict['mail'],dict['qq'],dict['note'])          
    phone = (dict['phone'],)
    try:
        cursor.execute("delete from student where phone = ?;",phone)
        cursor.execute('''insert into student(name,address,phone,wechat,mail,qq,note) 
                       values(?,?,?,?,?,?,?);''',values)
        cursor.close()
        conn.commit()
        conn.close()
        return jsonify({'finished': 'true'})
    except:
        print("Error: unable to change data")
        return jsonify({'finished': 'false'})

@application.route('/check', methods = ['POST'])                                           # 查看学生信息 
def check_stu():
    conn = sqlite3.connect('test.db')
    cursor = conn.cursor()
    listdata = []
    data = request.get_data()
    data = data.decode("utf-8")
    dict = json.loads(data) 
    phone = (dict['phone'],)
    try:
        cursor.execute("select* from student where phone = ?",phone)
        results = cursor.fetchall()
        for row in results:
            one_record = {}
            one_record['name'] = row[0]
            one_record['address'] = row[1]
            one_record['phone'] = row[2]
            one_record['wechat'] = row[3]
            one_record['mail'] = row[4]
            one_record['qq'] = row[5]
            one_record['note'] = row[6]
            listdata.append(one_record)

        cursor.close()
        conn.commit()
        conn.close()
        return tojson(listdata)        
    except:
        print("Error: unable to check data")
        return jsonify({'finished': 'false'})

if __name__ == '__main__':
    application.run()