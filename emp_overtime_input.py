import os
import sys
# import warnings
# import time

from PyQt5.QtWidgets import *
from PyQt5.QtCore import Qt, QSize, QDate, QTime
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
emp_overtime_input_window= uic.loadUiType(resource_path("./ui/emp_overtime_input.ui"))[0] # Window 사용시 ui 주소
# dial_window= uic.loadUiType(resource_path("C:\\myproject\\python project\\overtime\\popup_dept_info.ui"))[0] # Window 사용시 ui 주소

#화면을 띄우는데 사용되는 Class 선언
class MainWindow(QWidget, emp_overtime_input_window) :
    def __init__(self) :
        super().__init__()
        self.setupUi(self)
        self.setWindowTitle("사원 잔업시간 입력")
        self.slots()
        self.date.setDate(QDate.currentDate())
        # self.date = self.date_edit.date().toString("yyyyMMdd")
        self.time_start.setTime(QTime(18, 00)) 
        self.time_end.setTime(QTime(18, 00)) 
        self.setFixedSize(QSize(1079,823))

    def slots(self):
        # self.btn_search.clicked.connect(self.make_data)
        self.btn_close.clicked.connect(self.window_close)
        self.btn_select_dept.clicked.connect(self.popup_dept_info)
        self.btn_select_emp.clicked.connect(self.popup_emp_info)
        self.btn_input.clicked.connect(self.input_data)
        self.btn_delete.clicked.connect(self.delete_rows)
        self.btn_save.clicked.connect(self.upload)
        self.time_start.timeChanged.connect(self.calculate_overtime)
        self.time_end.timeChanged.connect(self.calculate_overtime)
        self.txt_dept_name.textChanged.connect(self.clear_empinfo)
        # self.btn_clear.clicked.connect(self.clear)

    # def set_date(self):
    #     date = self.date_select.date()
    #     self.txt_date.setText(date.toString("yyyy-MM"))

    def clear(self):        
        self.tbl_info.setRowCount(0) # clear()는 행은 그대로 내용만 삭제, 행을 "0" 호출 한다.

        self.txt_dept_id.setText("")
        self.txt_dept_name.setText("")
        self.txt_emp_id.setText("")
        self.txt_emp_name.setText("")

    def clear_empinfo(self):
        self.txt_emp_id.setText("")
        self.txt_emp_name.setText("")

    def input_data(self):

        overtime_date = self.date.date()
        overtime_date = overtime_date.toString("yyyy-MM-dd")

        dept_id = self.txt_dept_id.text()
        dept_name = self.txt_dept_name.text()
        emp_id = self.txt_emp_id.text()
        emp_name = self.txt_emp_name.text()
        overtime = self.txt_overtime.text()
        from_time = self.time_start.time().toString("hh:mm")
        to_time = self.time_end.time().toString("hh:mm")
        detail = self.txt_detail.text()
        note = self.txt_note.text()

        list = [dept_id, dept_name, emp_id, emp_name, overtime_date, from_time, to_time, str(overtime), detail, note]
        title = ["부서ID", "부서명", "사번", "사원명", "잔업일자", "시작시간", "종료시간", "잔업시간", "작업내용", "비고"]
        
        # 필수 입력값 
        list_must = [dept_id, dept_name, emp_id, emp_name, overtime_date, str(overtime)]
        title_must = ["부서ID", "부서명", "사번", "사원명", "잔업일자", "잔업시간"]

        # 필수 입력값 중 하나라도 ""이면 처리 중단
        for i, value in enumerate(list_must, start=1):
            if value == "":
                self.msg_box("입력오류", f"{title_must[i-1]} 값이 누락 됐습니다.")
                return

        # 잔업시간이 숫자(float)가 아닌경우 처리 중단
        try:
            float(overtime)
        except:
            self.msg_box("입력오류", "잔업시간 값이 숫자가 아닙니다.")
            return

        row_count = self.tbl_info.rowCount()
        col_count = len(title)
        
        self.tbl_info.insertRow(row_count)
        self.tbl_info.setColumnCount(col_count)
        self.tbl_info.setHorizontalHeaderLabels(title)

        for i in range(col_count):
            self.tbl_info.setItem(row_count, i, QTableWidgetItem(str(list[i])))
            self.tbl_info.item(row_count, i).setTextAlignment(Qt.AlignCenter | Qt.AlignVCenter)

            # self.txt_dept_id.setText("")
            # self.txt_dept_name.setText("")
            # self.txt_emp_id.setText("")
            # self.txt_emp_name.setText("")
            # self.txt_overtime.setText("")
            # self.txt_from_time.setText("")
            # self.txt_to_time.setText("")
            # self.txt_detail.setText("")
            # self.txt_note.setText("")
    
    # 입력한 시작/종료 시간으로 잔업시간 계산
    def calculate_overtime(self):
        time_1 = self.time_start.time()
        time_2 = self.time_end.time()

        # 시간 차이 계산
        secs = time_1.secsTo(time_2)
        hours, remainder = divmod(abs(secs), 3600)
        # minutes, seconds = divmod(remainder, 60)
        result = hours + round((remainder/3600), 1)
        self.txt_overtime.setText(str(result)) 
        # 가운데 정렬       
        self.txt_overtime.setAlignment(Qt.AlignCenter)

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

    def upload(self):
        # 현재 테이블 데이터(수정, 삭제 될 수 있다.)
        rows = self.tbl_info.rowCount()
        cols = self.tbl_info.columnCount()

        list = [] # 최종적으로 사용할 리스트는 for문 밖에 선언
        for i in range(rows):
            list_1 = []
            for j in range(cols):
                data = self.tbl_info.item(i,j)
                list_1.append(data.text())
            list.append(list_1)

        from db.db_insert import Insert
        data_insert = Insert()
        result = data_insert.insert_overtime(list)

        self.msg_box(result[0], result[1])
        self.tbl_info.setRowCount(0)

        self.txt_dept_id.setText("")
        self.txt_dept_name.setText("")
        self.txt_emp_id.setText("")
        self.txt_emp_name.setText("")
        self.txt_detail.setText("")
        self.txt_note.setText("")

        self.time_start.setTime(QTime(18, 00)) 
        self.time_end.setTime(QTime(18, 00)) 

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
        input_dialog = EmpWindow(arg_1) ##   <-----중요 포인트

        if input_dialog.exec_():
            value = input_dialog.get_input_value()

        try:
            self.txt_emp_id.setText(value[2].text())
            self.txt_emp_name.setText(value[3].text())
        except:
            return
        
    def get_dept_id(self):
        return self.dept_id
        
    def msg_box(self, arg_1, arg_2):
        msg = QMessageBox()
        msg.setWindowTitle(arg_1)               # 제목설정
        msg.setText(arg_2)                          # 내용설정
        msg.exec_()                                 # 메세지박스 실행

    def window_close(self):
        self.close()

if __name__ == "__main__" :
    app = QApplication(sys.argv)
    try:
        myWindow = MainWindow()
        myWindow.show()
        app.exec_()
    except Exception as e:
        msg = QMessageBox()
        msg.setWindowTitle("Error")               # 제목설정
        msg.setText(str(e))                          # 내용설정
        msg.exec_()  