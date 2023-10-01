from flask import Flask, render_template, request, after_this_request, jsonify, send_from_directory
import os
import tkinter as tk
from tkinter import messagebox
import threading
from openpyxl.styles import NamedStyle
import tkinter as tk
import openpyxl
import addlec
import addmon
import manager as mn
import Guilec

"""
    # Khởi tạo lớp Flask, tương tự với Tkinter
"""
class GuischeWeb(Flask):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        
        self.cansave = False
        self.path = os.getcwd() + "\\alpha.xlsx"
        self.checktrung = False
        self.list_copy = []
        
        self.upload_path = 'uploads'
        self.file_path = os.path.abspath(os.path.join(self.upload_path, 'tkb.xlsx'))
    
    def open(self):
        if self.file_path:
            print("Thư mục được chọn:", self.file_path)

            self.list_ = mn.readfile(self.file_path)
            for i in self.list_:
                self.list_copy.append(i)

            # mn.lopghep(self.list_)
            data = self.load_data(self.list_)
            self.cansave = True
            
            return jsonify(data)
        else:
            print("Không có thư mục nào được chọn.")
            
    def load_data(self, data):
        # data dạng đối tượng rồi mới chuyển thành dạng list
        list_data = mn.list_print(data)
            
        return list_data
    
    def xepgiangvien(self):
        try:
            list_xep = mn.readfile(self.file_path)
            self.checktrung = True
            # mn.lopghep(list_xep)
            mn.xepgiangvien(list_xep, "Sheet1")
            data = self.load_data(list_xep)
            self.list_ = list_xep
            filtered_teachers = [subarray[7] for subarray in data if len(subarray) > 7]
            return jsonify(filtered_teachers)
        except Exception as error:
            print('Error:', error)
            return jsonify('Error')
        
    def checkloptrung(self):
        if self.checktrung:
            mn.checktrung(self.list_)
            data = self.load_data(self.list_)
            
            filtered_duplications = [subarray[9] for subarray in data if len(subarray) > 9]
            return jsonify(filtered_duplications)
        else:
            print('Error')
            return jsonify('Error')


"""
    # Khởi tạo ứng dụng Website
"""
app = GuischeWeb(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    file = request.files['file']
    if file:
        filename = file.filename

        if not os.path.exists(app.upload_path):
            os.makedirs(app.upload_path)
        file.save(os.path.join(app.upload_path, filename))
        app.file_path = os.path.abspath(os.path.join(app.upload_path, filename))
        return app.open()
    else:
        return 'No file provided.', 400

# @app.route('/uploads/<filename>')
# def uploaded_file(filename):
#     return send_from_directory(app.file_path, filename)

@app.route('/arrange')
def arrange():
    if app.file_path:
        return app.xepgiangvien()
    else:
        return 'No file provided.', 400
    
@app.route('/check-duplications')
def checkDuplications():
    if app.checktrung:
        return app.checkloptrung()
    else:
        return 'No Permission.', 400

if __name__ == '__main__':
    app.run(debug=True)