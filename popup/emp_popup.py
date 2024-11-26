import os
import sys
# import warnings
# import time
from PyQt5.QtWidgets import *
from PyQt5.QtCore import Qt, QSize
from PyQt5 import uic

# 절대경로를 상대경로로 변경 하는 함수
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

#UI파일 연결
# main_window= uic.loadUiType(resource_path("/Users/black/projects/make_erp/main_window.ui"))[0] # Mac 사용시 ui 주소
emp_window = uic.loadUiType(resource_path("./ui/emp_window.ui"))[0]

class EmpWindow(QDialog, emp_window):
    def __init__(self, arg_1) :
        super().__init__()
        self.setupUi(self)
        self.setWindowTitle("사원선택")
        self.slots()

        self.dept_id = arg_1

        self.make_table()
        self.setFixedSize(QSize(540, 540))

    def slots(self):
        ### 다이알로그 시그널 생성기 반드시!!!!!!!! 필요. 없으면 작동 안 함############
        self.btn_confirm.clicked.connect(self.accept) # Close the dialog when OK is clicked 
        self.tbl_info.cellDoubleClicked.connect(self.accept)

    def make_table(self):
        from db.db_select import Select
        select = Select()
        select_emp = select.select_employee(self.dept_id)

        self.tbl_info.setRowCount(0) # clear()는 행은 그대로 내용만 삭제, 행을 "0" 호출 한다.

        if select_emp is None:
            num = 0
        else:
            num = len(select_emp)

        column_title = ["부서아이디", "부서명", "사원번호", "사원명"]
        col = len(column_title)

        self.tbl_info.setRowCount(num)
        self.tbl_info.setColumnCount(col)
        self.tbl_info.setHorizontalHeaderLabels(column_title)

        for i in range(num):
            for j in range(col): # 아니면 10개
                self.tbl_info.setItem(i, j, QTableWidgetItem(select_emp[i][j]))
                self.tbl_info.item(i, j).setTextAlignment(Qt.AlignCenter | Qt.AlignVCenter)  

        # 컨텐츠의 길이에 맞추어 컬럼의 길이를 자동으로 조절
        ################################################################
        table = self.tbl_info
        # header = table.horizontalHeader()
        table.setColumnWidth(0, int(table.width() * 0.25))
        table.setColumnWidth(1, int(table.width() * 0.25))
        table.setColumnWidth(2, int(table.width() * 0.25))
        table.setColumnWidth(3, int(table.width() * 0.25))

        # for i in range(col):
        #     header.setSectionResizeMode(i, QHeaderView.ResizeToContents)
        ################################################################

        # 테이블의 길이에 맞추어 컬럼 길이를 균등하게 확장
        self.tbl_info.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
    
    def get_input_value(self):
        list = self.tbl_info.selectedItems()
        return list

    def msg_box(self, arg_1, arg_2):
        msg = QMessageBox()
        msg.setWindowTitle(arg_1)               # 제목설정
        msg.setText(arg_2)                          # 내용설정
        msg.exec_()                                 # 메세지박스 실행

    def window_close(self):
        self.close()

if __name__ == "__main__" :
    app = QApplication(sys.argv)
    myWindow = EmpWindow()
    myWindow.show()
    app.exec_()