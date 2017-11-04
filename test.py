#coding:utf-8
import jsonify
import xlwt
from flask import Flask, request
from flask_script import Manager 
import sqlite3
import json
app = Flask(__name__)

@app.route('/export')                                                                    # 将数据库信息转换为excel
def export():                                                                            
    conn = sqlite3.connect('test.db')                                                    # 连接数据库
    cursor = conn.cursor()
    count = cursor.execute("select * from student")                                      # 执行sql语句

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

    workbook.save('demo.xls')

    cursor.close()                                                                       # 关闭数据库连接
    conn.commit()
    conn.close()
    return jsonify({'finished': True})

def tojson(listdata):                                                                    # 返回json格式数据
    senddata = {}
    senddata['data'] = listdata
    print(json.dumps(senddata, ensure_ascii=False))
    return  json.dumps(senddata, ensure_ascii=False)

@app.route('/show')                                                                      # 显示所有学生信息
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
    except:
        print("Error: unable to fecth data")
    
    cursor.close()
    conn.commit()
    conn.close()

    return tojson(listdata)
    
@app.route('/add', methods = ['POST','GET'])                                             # 添加学生信息
def add_stu():
    conn = sqlite3.connect('test.db')
    cursor = conn.cursor()
    if request.method == 'POST':                                                              # 接受前端发送的信息
        my_name = request.form.get('name', '')
        my_address = request.form.get('address','')
        my_phone = request.form.get('phone','')
        my_wechat = request.form.get('wechat','')
        my_mail = request.form.get('mail','')
        my_qq = request.form.get('qq','')
        my_note = request.form.get('note','')
    else:
        my_name = request.args.get('name', '')
        my_address = request.args.get('address','')
        my_phone = request.args.get('phone','')
        my_wechat = request.args.get('wechat','')
        my_mail = request.args.get('mail','')
        my_qq = request.args.get('qq','')
        my_note = request.args.get('note','')
    try:
        cursor.execute("insert into student(name,address,phone,wechat,mail,qq,note) values('"+
                       dform['my_name']+"','"+dform['my_address']+"','"+dform['my_phone']+"','"+dform['my_wechat']+
                       "','"+dform['my_mail']+"','"+dform['my_qq']+"','"+dform['my_note']+"');")
    except:
        print("Error: unable to insert data")

    cursor.close()
    conn.commit()
    conn.close()
    return jsonify({'finished': True})

@app.route('/delete', methods = ['POST','GET'])                                          # 删除学生信息                 
def delete_stu():
    conn = sqlite3.connect('test.db')
    cursor = conn.cursor()
    if request.method == 'POST':
        phone = request.form.get('phone', '')
    else:
        phone = request.args.get('phone', '')
    try:
        cursor.execute("delete from student where phone = '"+phone+"';")
    except:
        print("Error: unable to delete data")
    cursor.close()
    conn.commit()
    conn.close()
    return jsonify({'finished': True})

@app.route('/change', methods = ['POST','GET'])                                          # 更改学生信息
def change_stu():
    conn = sqlite3.connect('test.db')
    cursor = conn.cursor()
    if request.method == 'POST':
        phone = request.form.get('phone', '')
    else:
        phone = request.args.get('phone', '')
    try:
        cursor.execute("delete from student where phone = '"+dform['phone']+"';")
        cursor.execute("insert into student(name,address,phone,wechat,mail,qq,note) values('"+
                       dform['my_name']+"','"+dform['my_address']+"','"+dform['my_phone']+"','"+dform['my_wechat']+
                       "','"+dform['my_mail']+"','"+dform['my_qq']+"','"+dform['my_note']+"');")
    except:
        print("Error: unable to change data")
    cursor.close()
    conn.commit()
    conn.close()
    return jsonify({'finished': True})

@app.route('/check', methods = ['POST','GET'])                                           # 查看学生信息 
def check_stu():
    conn = sqlite3.connect('test.db')
    cursor = conn.cursor()
    if request.method == 'POST':
        phone = request.form.get('phone', '')
    else:
        phone = request.args.get('phone', '')
    try:
        cursor.execute("select* from student where phone = "+dform['my_phone'])
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
            
    except:
        print("Error: unable to check data")
    
    cursor.close()
    conn.commit()
    conn.close()

    return tojson(listdata)

if __name__ == '__main__':
    app.run(debug=True)