import os
import sys
# import warnings
# import time

from PyQt5.QtWidgets import *
from PyQt5.QtCore import Qt, QDate, QSize
from PyQt5 import uic
import openpyxl
from openpyxl.styles import Alignment
from datetime import datetime

# 절대경로를 상대경로로 변경 하는 함수
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

#UI파일 연결
main_window= uic.loadUiType(resource_path("./ui/input_prod_order.ui"))[0] # Window 사용시 ui 주소
# main_window= uic.loadUiType(resource_path("/Users/black/projects/make_erp/main_window.ui"))[0] # Mac 사용시 ui 주소

#화면을 띄우는데 사용되는 Class 선언
class MainWindow(QWidget, main_window) :
    def __init__(self) :
        super().__init__()
        self.setupUi(self)
        self.setWindowTitle("생산오더 입력")

        self.layout_setting()
        self.slots()


    def layout_setting(self):
        # 필수 레이아웃
        layout_essential = QHBoxLayout()
        layout_essential.addWidget(self.lbl_order_type)
        layout_essential.addWidget(self.comb_order_type)
        layout_essential.addWidget(self.lbl_dept_origin)
        layout_essential.addWidget(self.comb_dept_origin)
        layout_essential.addWidget(self.lbl_p_order_id)
        layout_essential.addWidget(self.txt_p_order_id) 

        # 아이템_1 레이아웃
        layout_item_1 = QHBoxLayout()
        layout_item_1.addWidget(self.lbl_item_id)
        layout_item_1.addWidget(self.txt_item_id)
        layout_item_1.addWidget(self.lbl_item_name)
        layout_item_1.addWidget(self.txt_item_name)
        layout_item_1.addWidget(self.lbl_item_qty)
        layout_item_1.addWidget(self.txt_item_qty)

        # 아이템_2 레이아웃
        layout_item_2 = QGridLayout()
        layout_item_2.addWidget(self.lbl_order_min, 0, 0, Qt.AlignLeft)
        layout_item_2.addWidget(self.txt_order_min, 0, 1, Qt.AlignLeft)
        layout_item_2.addWidget(self.lbl_status, 0, 2, Qt.AlignLeft)
        layout_item_2.addWidget(self.comb_status, 0, 3, Qt.AlignLeft)
        layout_item_2.addWidget(self.lbl_s_order_id, 1, 0, Qt.AlignLeft)
        layout_item_2.addWidget(self.txt_s_order_id, 1, 1, Qt.AlignLeft)
        layout_item_2.addWidget(self.lbl_s_date, 1, 2, Qt.AlignLeft)
        layout_item_2.addWidget(self.date_s_date, 1, 3, Qt.AlignLeft)
        layout_item_2.addWidget(self.lbl_p_dept_id, 2, 0, Qt.AlignLeft)
        layout_item_2.addWidget(self.comb_p_dept_id, 2, 1, Qt.AlignLeft)

        # 실행 레이아웃
        layout_exec = QHBoxLayout()
        layout_exec.addWidget(self.btn_save)
        layout_exec.addWidget(self.btn_close)

        # 기타 위젯 선언
        self.lbl_title_1 = QLabel("1. 필수 입력항목")
        self.lbl_title_1.setFixedHeight(20)

        self.lbl_title_2 = QLabel("2. 선택 입력항목")
        self.lbl_title_2.setFixedHeight(20)

        self.line_1 = QFrame()
        self.line_1.setFrameShape(QFrame.HLine)  # 수평선
        self.line_1.setFrameShadow(QFrame.Sunken)

        self.line_2 = QFrame()
        self.line_2.setFrameShape(QFrame.HLine)  # 수평선
        self.line_2.setFrameShadow(QFrame.Sunken)


        # # 전체 레이아웃
        main_layout = QVBoxLayout()
        main_layout.addWidget(self.lbl_title_1)
        main_layout.addLayout(layout_essential)
        main_layout.addLayout(layout_item_1)
        main_layout.addWidget(self.line_1)
        main_layout.addWidget(self.lbl_title_2)
        main_layout.addLayout(layout_item_2)
        main_layout.addWidget(self.line_2)
        main_layout.addLayout(layout_exec)
        main_layout.addStretch() 
        # main_layout.addWidget(self.label_1)
        # main_layout.addLayout(self.layout_essential)
        # main_layout.addLayout(self.layout_item_1)
        # main_layout.addWidget(self.line_1)
        # main_layout.addWidget(self.label_2)
        # main_layout.addLayout(self.layout_item_2)
        # main_layout.addWidget(self.line_2)
        # main_layout.addWidget(self.btn_save)
        # main_layout.addWidget(self.btn_close)

        self.setFixedSize(670, 300)


        self.setLayout(main_layout)

        # 콤보박스 설정
        items_status = ['', '릴리스됨', '시작됨', '중지됨', '종료됨']
        self.comb_status.addItems(items_status)
        self.comb_status.setCurrentIndex(1)

        items_order_type = ['', '생산오더', '분해오더', '재작업오더']
        self.comb_order_type.addItems(items_order_type)

        items_dept_origin = ['', '생산본부', '영업부']
        self.comb_dept_origin.addItems(items_dept_origin)

        items_p_dept_name = ['', '제조1파트', '제조2파트', '제조3파트', '제조4파트', '제조1,4파트']
        self.comb_p_dept_id.addItems(items_p_dept_name)

        # 텍스트위젯 설정
        self.txt_order_min.setText("0")        

        # 현재시간 설정
        self.set_date()

    def slots(self):
        self.btn_save.clicked.connect(self.get_args)
        self.btn_close.clicked.connect(self.window_close)

    def set_date(self):
        self.date_s_date.setDate(QDate.currentDate())

    def clear(self):     
        self.comb_order_type.setCurrentIndex(0)
        self.comb_dept_origin.setCurrentIndex(0)
        self.txt_p_order_id.setText("")
        self.txt_item_id.setText("")
        self.txt_item_name.setText("")
        self.txt_item_qty.setText("")
        self.txt_order_min.setText("0")
        self.comb_status.setCurrentIndex(1)
        self.txt_s_order_id.setText("")
        self.date_s_date.setDate(QDate.currentDate())   
        self.comb_p_dept_id.setCurrentIndex(0)

    def get_args(self):

        #  DB에 입력할 값 수집
        order_type = self.comb_order_type.currentText()
        dept_origin = self.comb_dept_origin.currentText()
        p_order_id = self.txt_p_order_id.text()
        item_id = self.txt_item_id.text()
        item_name = self.txt_item_name.text()
        item_qty = self.txt_item_qty.text()
        order_min = self.txt_order_min.text()
        status = self.comb_status.currentText()
        s_order_id = self.txt_s_order_id.text()
        s_date = self.date_s_date.dateTime().toString("yyyy-MM-dd hh:mm:ss")
        
        # 생산 할당 부서는 부서이름과 매칭되는 부서 코드를 가져온다.
        items_p_dept_id = ['', '1100', '1200', '1300', '1400', '1410']
        index = self.comb_p_dept_id.currentIndex()
        p_dept_id = items_p_dept_id[index]

        if order_type == "" or dept_origin == "" or p_order_id == "" or item_id == "" or item_name == "" or item_qty == "" or p_order_id == "" :
            self.msg_box("Warning", "필수 입력값을 입력하세요.")
            return
        
        # INSERT INTO production_upload (p_order_id, item_id, item_name, status, p_dept_id, s_order_id, s_date, item_qty, order_min, order_type, dept_origin)

        arr_1 = [p_order_id, item_id, item_name, status,  p_dept_id, s_order_id, s_date, item_qty, order_min, order_type, dept_origin]
        self.save_data(arr_1)

    def save_data(self, arr_1):
           
        from db.db_insert import Insert
        insert = Insert()

        try:
            result = insert.input_prod_info(arr_1)
            self.msg_box(result[0], result[1])
        except Exception as e:
                self.msg_box("Program Error", str(e))
                return

    def msg_box(self, arg_1, arg_2):
        msg = QMessageBox()
        msg.setWindowTitle(arg_1)               # 제목설정
        msg.setText(arg_2)                          # 내용설정
        msg.exec_()                                 # 메세지박스 실행

    def window_close(self):
        self.close()

if __name__ == "__main__" :
    app = QApplication(sys.argv)
    myWindow = MainWindow()
    myWindow.show()
    app.exec_()