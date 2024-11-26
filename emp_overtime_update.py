import os
import sys
# import warnings
# import time

from PyQt5.QtWidgets import *
from PyQt5.QtCore import Qt, QSize, QDate, QTime
from PyQt5 import uic, QtWidgets, QtCore

# 절대경로를 상대경로로 변경 하는 함수
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

#UI파일 연결
# main_window= uic.loadUiType(resource_path("/Users/black/projects/make_erp/main_window.ui"))[0] # Mac 사용시 ui 주소
emp_overtime_update_window= uic.loadUiType(resource_path("./ui/emp_overtime_update.ui"))[0] # Window 사용시 ui 주소

# dial_window= uic.loadUiType(resource_path("C:\\myproject\\python project\\overtime\\popup_dept_info.ui"))[0] # Window 사용시 ui 주소

#화면을 띄우는데 사용되는 Class 선언
class MainWindow(QWidget, emp_overtime_update_window) :
    def __init__(self) :
        super().__init__()
        self.setupUi(self)
        self.setWindowTitle("사원 잔업시간 수정")
        self.slots()
        self.date_from.setDate(QDate.currentDate())
        self.date_to.setDate(QDate.currentDate())

        self.time_start.setTime(QtCore.QTime(18, 00)) 
        self.time_end.setTime(QtCore.QTime(18, 00)) 

        # self.date = self.date_edit.date().toString("yyyyMMdd")
        self.setFixedSize(QSize(1079,823))

    def slots(self):
        # self.btn_search.clicked.connect(self.make_data)
        self.btn_close.clicked.connect(self.window_close)
        self.btn_select_dept.clicked.connect(self.popup_dept_info)
        self.btn_select_dept.clicked.connect(self.clear_txt)
        self.btn_select_emp.clicked.connect(self.popup_emp_info)
        self.btn_select.clicked.connect(self.search_overtime_info)
        self.btn_delete.clicked.connect(self.delete_overtime_info)
        self.btn_update.clicked.connect(self.update_overtime_info)
        self.tbl_info.cellClicked.connect(self.select_info)       
        self.time_start.timeChanged.connect(self.calculate_overtime)
        self.time_end.timeChanged.connect(self.calculate_overtime)
        # self.txt_dept_id.textChanged().connect(self.clear_txt)
        # self.btn_save.clicked.connect(self.upload)
        # self.btn_input.clicked.connect(self.input_data)
        # self.btn_clear.clicked.connect(self.clear)

    # def set_date(self):
    #     date = self.date_select.date()
    #     self.txt_date.setText(date.toString("yyyy-MM"))
    def clear_txt(self):
        self.txt_emp_id.setText("")
        self.txt_emp_name.setText("")
        self.txt_overtime.setText("")
        self.time_start.setTime(QtCore.QTime(18, 00)) 
        self.time_end.setTime(QtCore.QTime(18, 00)) 
        self.txt_detail.setText("")
        self.txt_note.setText("")

    def clear(self):        
        self.tbl_info.setRowCount(0) # clear()는 행은 그대로 내용만 삭제, 행을 "0" 호출 한다.

        self.txt_dept_id.setText("")
        self.txt_dept_name.setText("")
        self.txt_emp_id.setText("")
        self.txt_emp_name.setText("")

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

    def search_overtime_info(self):
        date_from = self.date_from.date().toString("yyyy-MM-dd")
        date_to = self.date_to.date().toString("yyyy-MM-dd")

        dept_id = self.txt_dept_id.toPlainText()
        if dept_id == "":
            dept_id = '%%'
        else:
            dept_id

        emp_id = self.txt_emp_id.toPlainText()
        if emp_id == "":
            emp_id = '%%'
        else:
            emp_id

        arr = [date_from, date_to, dept_id, emp_id]

        from db.db_select import Select
        select = Select()
        result = select.update_overtime(arr)

        if result is None:
            return        
        else:
            title = ["ID", "부서아이디", "부서명", "사번", "이름", "날짜", "잔업시간", "시작시간", "종료시간", "작업내용", "비고"]
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

    def select_info(self):
        x = self.tbl_info.selectedIndexes() # 리스트로 선택된 행번호와 열번호가 x에 입력된다.
        row = x[0].row() #첫번째 선택된 행번호를 부르는 방법
        num_col = self.tbl_info.columnCount()

        list = []
        for i in range(num_col):
            value = self.tbl_info.item(row, i).text()
            list.append(value)

        date = list[5]
        year = int(date[0:4])
        month = int(date[5:7])
        day = int(date[8:10])

        d = QDate(year, month, day)

        self.date_from.setDate(d)
        self.date_to.setDate(d)

        # 시간 자르기 편집
        start = list[7]
        end = list[8]

        start_hh = int(start[0:2])
        start_mm = int(start[3:5])
        end_hh = int(end[0:2])
        end_mm = int(end[3:5])

        
        self.txt_id.setText(list[0])
        self.txt_dept_id.setText(list[1])
        self.txt_dept_name.setText(list[2])
        self.txt_emp_id.setText(list[3])
        self.txt_emp_name.setText(list[4])
        self.txt_overtime.setText(list[6])
        self.time_start.setTime(QTime(start_hh, start_mm))
        self.time_end.setTime(QTime(end_hh, end_mm))
        self.txt_detail.setText(list[9])
        self.txt_note.setText(list[10])

    # def delete_overtime_info(self):
    #     id = self.txt_id.toPlainText()
    #     dept_name = self.txt_dept_name.toPlainText()
    #     emp_name = self.txt_emp_name.toPlainText()

    #     option = QtWidgets.QMessageBox.question(self, "QMessageBox", f"{dept_name} {emp_name} 사원의 잔업정보를 삭제 하시겠습니까?", 
    #                                    QtWidgets.QMessageBox.No | QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.Cancel, QtWidgets.QMessageBox.Yes)
        
    #     if option == QtWidgets.QMessageBox.Cancel:
    #         return
    #     elif option == QtWidgets.QMessageBox.No:
    #         return
    #     elif option == QtWidgets.QMessageBox.Yes: 
            
    #         from db.db_delete import Delete
    #         delete = Delete()
    #         delete.delete_emp_overtime(id)

    #         self.tbl_info.setRowCount(0)            

    def update_overtime_info(self):
        now = QDate.currentDate()
        date_from = self.date_from.date()
        days = date_from.daysTo(now)

        if days > 30:
            self.msg_box("수정오류", "수정 가능 기간이 지났습니다.")
            return
        else:
            id = self.txt_id.toPlainText()
            dept_name = self.txt_dept_name.toPlainText()
            emp_name = self.txt_emp_name.toPlainText()

            overtime = self.txt_overtime.toPlainText()
            from_time = self.time_start.time().toString("hh:mm")
            to_time = self.time_end.time().toString("hh:mm")
            detail = self.txt_detail.toPlainText()
            note = self.txt_note.toPlainText()

            update_list = [overtime, from_time, to_time, detail, note, id]

            option = QtWidgets.QMessageBox.question(self, "QMessageBox", f"{dept_name} {emp_name} 사원의 잔업정보를 수정 하시겠습니까?", 
                                        QtWidgets.QMessageBox.No | QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.Cancel, QtWidgets.QMessageBox.Yes)
            
            if option == QtWidgets.QMessageBox.Cancel:
                return
            elif option == QtWidgets.QMessageBox.No:
                return
            elif option == QtWidgets.QMessageBox.Yes: 

                from db.db_update import Update
                update = Update()
                update.update_emp_overtime(update_list)
                self.tbl_info.setRowCount(0)

                self.date_from.setDate(QDate.currentDate())
                self.date_to.setDate(QDate.currentDate())

                self.txt_dept_id.setText("")
                self.txt_dept_name.setText("")
                self.txt_emp_id.setText("")
                self.txt_emp_name.setText("")
                # self.txt_overtime.setText("")
                self.time_start.setTime(QTime(18, 00)) 
                self.time_end.setTime(QTime(18, 00)) 
                self.txt_detail.setText("")
                self.txt_note.setText("")

                from db.db_select import Select
                select = Select()
                result = select.update_overtime_id(id)

                title = ["ID", "부서아이디", "부서명", "사번", "이름", "날짜", "잔업시간", "시작시간", "종료시간", "작업내용", "비고"]
                self.make_table(len(result), result, title)

    def delete_overtime_info(self):
        now = QDate.currentDate()
        date_from = self.date_from.date()
        days = date_from.daysTo(now)

        if days > 30:
            self.msg_box("삭제오류", "삭제 가능 기간이 지났습니다.")
            return
        else:
            id = self.txt_id.toPlainText()
            dept_name = self.txt_dept_name.toPlainText()
            emp_name = self.txt_emp_name.toPlainText()

            option = QtWidgets.QMessageBox.question(self, "QMessageBox", f"{dept_name} {emp_name} 사원의 잔업정보를 삭제 하시겠습니까?", 
                                        QtWidgets.QMessageBox.No | QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.Cancel, QtWidgets.QMessageBox.Yes)
            
            if option == QtWidgets.QMessageBox.Cancel:
                return
            elif option == QtWidgets.QMessageBox.No:
                return
            elif option == QtWidgets.QMessageBox.Yes: 

                from db.db_delete import Delete
                delete = Delete()
                delete.delete_emp_overtime(id)
                self.tbl_info.setRowCount(0)               
                
                # self.search_overtime_info()

                self.date_from.setDate(QDate.currentDate())
                self.date_to.setDate(QDate.currentDate())

                self.txt_dept_id.setText("")
                self.txt_dept_name.setText("")
                self.txt_emp_id.setText("")
                self.txt_emp_name.setText("")
                # self.txt_overtime.setText("")
                self.txt_detail.setText("")
                self.txt_note.setText("")

                self.time_start.setTime(QTime(18, 00)) 
                self.time_end.setTime(QTime(18, 00)) 

    def input_data(self):

        overtime_date = self.date.date()
        overtime_date = overtime_date.toString("yyyy-MM-dd")

        dept_id = self.txt_dept_id.toPlainText()
        dept_name = self.txt_dept_name.toPlainText()
        emp_id = self.txt_emp_id.toPlainText()
        emp_name = self.txt_emp_name.toPlainText()
        overtime = self.txt_overtime.toPlainText()
        from_time = self.txt_from_time.toPlainText()
        to_time = self.txt_to_time.toPlainText()
        detail = self.txt_detail.toPlainText()
        note = self.txt_note.toPlainText()

        list = [dept_id, dept_name, emp_id, emp_name, overtime_date, str(overtime), from_time, to_time, detail, note]
        title = ["부서ID", "부서명", "사번", "사원명", "잔업일자", "잔업시간", "시작시간", "종료시간", "작업내용", "비고"]
        
        # 필수 입력값 
        list_must = [dept_id, dept_name, emp_id, emp_name, overtime_date, str(overtime)]
        title_must = ["부서ID", "부서명", "사번", "사원명", "잔업일자", "잔업시간"]

        # 필수 입력값 중 하나라도 ""이면 처리 중단
        for i, value in enumerate(list_must, start=1):
            if value == "":
                self.msg_box("입력오류", f"{title_must[i-1]} 값이 누락 됐습니다.")
                return

        row_count = self.tbl_info.rowCount()
        col_count = len(title)
        
        self.tbl_info.insertRow(row_count)
        self.tbl_info.setColumnCount(col_count)
        self.tbl_info.setHorizontalHeaderLabels(title)

        for i in range(col_count):
            self.tbl_info.setItem(row_count, i, QTableWidgetItem(str(list[i])))
            self.tbl_info.item(row_count, i).setTextAlignment(Qt.AlignCenter | Qt.AlignVCenter)

            self.txt_dept_id.setText("")
            self.txt_dept_name.setText("")
            self.txt_emp_id.setText("")
            self.txt_emp_name.setText("")
            self.txt_overtime.setText("")
            self.txt_from_time.setText("")
            self.txt_to_time.setText("")
            self.txt_detail.setText("")
            self.txt_note.setText("")
        
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
        arg_1 = self.txt_dept_id.toPlainText()
        
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
    myWindow = MainWindow()
    myWindow.show()
    app.exec_()