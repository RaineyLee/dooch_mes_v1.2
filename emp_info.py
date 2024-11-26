import os
import sys

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
emp_window = uic.loadUiType(resource_path("./ui/emp_master.ui"))[0]
# dept_window = uic.loadUiType(resource_path("./ui/dept_window.ui"))[0]
# total_overtime= uic.loadUiType(resource_path("C:\\myproject\\python project\\overtime\\overtime_v1.1\\ui\\total_overtime.ui"))[0] # Window 사용시 ui 주소
# main_window= uic.loadUiType(resource_path("/Users/black/projects/make_erp/main_window.ui"))[0] # Mac 사용시 ui 주소

#화면을 띄우는데 사용되는 Class 선언
class MainWindow(QWidget, emp_window) :
    def __init__(self) :
        super().__init__()
        self.setupUi(self)
        self.setWindowTitle("인사정보 조회")
        self.slots()

        use = ['y', 'n']
        self.cmb_yn.addItems(use)

        self.setFixedSize(QSize(880, 600))
        
    def slots(self):
        self.btn_dept_search.clicked.connect(self.popup_dept_info)
        self.btn_emp_search.clicked.connect(self.popup_emp_info)
        self.btn_search.clicked.connect(self.make_data)
        self.btn_close.clicked.connect(self.window_close)
        self.txt_dept_name.textChanged.connect(self.clear_emp)
        # self.btn_clear.clicked.connect(self.clear)
        # self.btn_close.clicked.connect(self.close)
        # self.btn_download.clicked.connect(self.make_file)
        # self.btn_select_emp.clicked.connect(self.popup_emp_info)

    # def set_date(self):
    #     date = self.date_select.date()
    #     self.txt_date.setText(date.toString("yyyy-MM"))

    def clear(self):        
        self.tbl_info.setRowCount(0) # clear()는 행은 그대로 내용만 삭제, 행을 "0" 호출 한다.

        self.txt_dept_id.setText("")
        self.txt_dept_name.setText("")

    def clear_emp(self):
        self.txt_emp_id.setText("")
        self.txt_emp_name.setText("")

    def make_data(self):
        dept_id = self.txt_dept_id.text()
        emp_id = self.txt_emp_id.text()
        use = self.cmb_yn.currentText()

        if dept_id and emp_id:

            from db.db_select import Select
            select = Select()
            result = select.emp_info_dept_emp([emp_id, use])
            
            if result is None:
                return            
            else:
                title = ["부서ID", "부서명", "사번", "이름", "사용"]
                self.make_table(len(result), result, title)
        elif  dept_id:
           
            from db.db_select import Select
            select = Select()

            result = select.emp_info_dept([dept_id, use])
            if result is None:
                return            
            else:
                title = ["부서ID", "부서명", "사번", "이름", "사용"]
                self.make_table(len(result), result, title)
        elif dept_id == "":
           
            from db.db_select import Select
            select = Select()
            result = select.emp_info(use)

            if result is None:
                return
            else:
                title = ["부서ID", "부서명", "사번", "이름", "사용"]
                self.make_table(len(result), result, title)

    def make_table(self, num, arr_1, title):   
        self.tbl_info.setRowCount(0) # clear()는 행은 그대로 내용만 삭제, 행을 "0" 호출 한다.

        col = len(title)

        self.tbl_info.setRowCount(num)
        self.tbl_info.setColumnCount(col)
        self.tbl_info.setHorizontalHeaderLabels(title)

        for i in range(num):
            for j in range(col): # 아니면 10개
                self.tbl_info.setItem(i, j, QTableWidgetItem(str(arr_1[i][j])))
                self.tbl_info.item(i, j).setTextAlignment(Qt.AlignCenter | Qt.AlignVCenter)     

        # 컨텐츠의 길이에 맞추어 컬럼의 길이를 자동으로 조절
        ################################################################
        table = self.tbl_info
        header = table.horizontalHeader()

        for i in range(col):
            header.setSectionResizeMode(i, QHeaderView.ResizeToContents)
        ################################################################

        # 테이블의 길이에 맞추어 컬럼 길이를 균등하게 확장
        self.tbl_info.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

    
    # 테이블 선택범위 삭제
    def delete_rows(self):
        indexes = []
        rows = []

        for idx in self.tbl_info.selectedItems():
            indexes.append(idx.row())

        for value in indexes:
            if value not in rows:
                rows.append(value)

        # 삭제시 오류 방지를 위해 아래서 부터 삭제(리버스 소팅)
        rows = sorted(rows, reverse=True)

        # 선택행 삭제
        for rowid in rows:
            self.tbl_info.removeRow(rowid)

    # 부서명 가져오기 팝업
    def popup_dept_info(self):
        from popup.dept_popup import DeptWindow
        input_dialog = DeptWindow()

        if input_dialog.exec_():
            value = input_dialog.get_input_value()

        try:
            self.txt_dept_id.setText(value[0].text())
            self.txt_dept_name.setText(value[1].text())
        except:
            return
        
    ### 다이알로그 창으로 값을 전달 할 때는 아규먼트를 보내 주는 방식으로 !!!!
    def popup_emp_info(self):
        arg_1 = self.txt_dept_id.text()

        from popup.emp_popup import EmpWindow
        input_dialog = EmpWindow(arg_1)##   <-----중요 포인트

        if input_dialog.exec_():
            value = input_dialog.get_input_value()

        try:
            self.txt_emp_id.setText(value[2].text())
            self.txt_emp_name.setText(value[3].text())
        except:
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