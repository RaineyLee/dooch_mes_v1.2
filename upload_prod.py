import os
import sys
# import warnings
# import time

from openpyxl import *
from PyQt5.QtWidgets import *
from PyQt5 import uic
from PyQt5.QtCore import Qt

# 절대경로를 상대경로로 변경 하는 함수
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

#UI파일 연결
# main_window= uic.loadUiType(resource_path("/Users/black/projects/make_erp/main_window.ui"))[0] # Mac 사용시 ui 주소
main_window= uic.loadUiType(resource_path("./ui/upload_prod.ui"))[0] # Window 사용시 ui 주소

# dial_window= uic.loadUiType(resource_path("C:\\myproject\\python project\\overtime\\popup_dept_info.ui"))[0] # Window 사용시 ui 주소

#화면을 띄우는데 사용되는 Class 선언
class MainWindow(QWidget, main_window) :
    def __init__(self) :
        super().__init__()
        self.setupUi(self)
        self.setWindowTitle("생산계획 업로드")
        self.slots()
        
        self.layout_setting()
        # self.slots()

    def layout_setting(self):
        # 버튼 레이아웃
        items_layout = QHBoxLayout()
        items_layout.addWidget(self.lbl_select_file)
        items_layout.addWidget(self.txt_select_file)
        items_layout.addWidget(self.btn_open)
        items_layout.addWidget(self.btn_select)

        # 실행 버튼 레이아웃
        exec_layout = QHBoxLayout()
        exec_layout.addWidget(self.btn_delete)
        exec_layout.addWidget(self.btn_upload)
        exec_layout.addWidget(self.btn_close)
        exec_layout.setAlignment(Qt.AlignRight)  # 오른쪽 정렬

        # 전체 레이아웃
        main_layout = QVBoxLayout()
        main_layout.addLayout(items_layout)  # 버튼 추가
        main_layout.addWidget(self.tbl_info)  # 테이블 추가
        main_layout.addLayout(exec_layout)

        self.setLayout(main_layout)

        # 현재시간 설정
        # self.set_date()

    def slots(self):
        self.btn_open.clicked.connect(self.file_open)
        self.btn_select.clicked.connect(self.make_data)
        self.btn_upload.clicked.connect(self.check_data)
        self.btn_close.clicked.connect(self.window_close)
        self.btn_delete.clicked.connect(self.delete_rows)
        # self.btn_select_dept.clicked.connect(self.popup_dept_info)

    # def set_date(self):
    #     self.date = self.date_edit.date().toString("yyyyMMdd")

    def file_open(self):
        fname = QFileDialog.getOpenFileName(parent=self, caption='Open file', directory='./excel/')

        if fname[0]:
            self.txt_select_file.setText(fname[0])
        else:
            self.txt_select_file.setText("")
            QMessageBox.about(self, 'Warning', '파일을 선택하지 않았습니다.')

    def make_data(self):
        file_name = self.txt_select_file.text()
        if file_name == "":
            return
        else:
            self.tbl_info.setRowCount(0) # clear()는 행은 그대로 내용만 삭제, 행을 "0" 호출 한다.

            from utils.make_data import Prodinfo
            make_data = Prodinfo(file_name)

            _list = make_data.excel_data()

            title = _list[1]
            data = _list[0]
                    
            self.make_table(len(data), data, title)

    def make_table(self, num, arr_1, arg_1):   
        self.tbl_info.setSortingEnabled(False)  # 정렬 비활성화 --> 정렬을 비활성화 하지 않으면 재/조회시 테이블이 꼬인다.
        self.tbl_info.setRowCount(0) # clear()는 행은 그대로 내용만 삭제, 행을 "0" 호출 한다.

        col = len(arg_1)

        self.tbl_info.setRowCount(num)
        self.tbl_info.setColumnCount(col)
        self.tbl_info.setHorizontalHeaderLabels(arg_1)

        for i in range(num):
            for j in range(col): # 아니면 10개를 적어도 된다.
                cell_value = arr_1[i][j]

                # NULL(None)을 공란으로 처리
                if cell_value == "None":
                    cell_value = ""
                
                # cell_value가 소수점인지 확인하고 정수로 변환
                if isinstance(cell_value, str) and cell_value.count('.') == 1:
                    try:
                        float_value = float(cell_value)
                        cell_value = int(float_value)
                    except ValueError:
                        pass                

                item = QTableWidgetItem(str(cell_value))
                self.tbl_info.setItem(i, j, item)
                # self.tbl_info.setItem(i, j, QTableWidgetItem(arr_1[i][j])) # 이렇게 해도 된다.

                 # 3번째 컬럼만 왼쪽 정렬
                if j == 2:  # 컬럼 인덱스 2 (3번째 컬럼)
                    item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                else:  # 나머지 컬럼은 중앙 정렬
                    item.setTextAlignment(Qt.AlignCenter | Qt.AlignVCenter) 

        # 컨텐츠의 길이에 맞추어 컬럼의 길이를 자동으로 조절
        ################################################################
        table = self.tbl_info
        header = table.horizontalHeader()

        # QSS 스타일 적용 (헤더 배경 색을 연한 회색으로 변경)
        table.setStyleSheet("""
            QHeaderView::section {
                background-color: lightgray;
                color: black;
                border: 1px solid #d6d6d6;
            }
        """)

        # 컬럼별 설정: 일부는 Interactive, 일부는 ResizeToContents
        for i in range(table.columnCount()):
            if i in [2, 5, 6]:  # 특정 컬럼은 길이에 맞추어 조정
                header.setSectionResizeMode(i, QHeaderView.ResizeToContents)
                
            else:  # 나머지 컬럼은 Interactive
                header.setSectionResizeMode(i, QHeaderView.Interactive)
        ################################################################

        # 정렬 기능 활성화
        self.tbl_info.setSortingEnabled(True)

        # 마지막 컬럼도 Stretch 비율로 포함
        header.setStretchLastSection(False)

       

        # 테이블의 길이에 맞추어 컬럼 길이를 균등하게 확장
        # self.tbl_info.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
    
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

    def check_data(self):
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

        arr_1 = []
        for i in list:
            arr_11 = (i[0])
            arr_1.append(str(arr_11))   
        arr_1 = tuple((arr_1))

        try:
            from db.db_check import Check
            check_info = Check()
            _check = check_info.check_prod_id(arr_1)

            if _check == True:
                self.upload(list)
            else:
                db_index = _check

                reply = QMessageBox.question(self, 'Message', '중복된 생산지시번호가 있습니다. 삭제하시겠습니까?', QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
                if reply == QMessageBox.Yes:
                    self.delete_rows_indexs(db_index)

        except Exception as e:
            self.msg_box("Error", str(e))
            return   
        
            # 테이블 선택범위 삭제
    def delete_rows_indexs(self, db_index):
        indexes = db_index
        rows = []

        for value in indexes:
            if value not in rows:
                rows.append(value)

        # 삭제시 오류 방지를 위해 아래서 부터 삭제(리버스 소팅)
        rows = sorted(rows, reverse=True)

        # 선택행 삭제
        for rowid in rows:
            self.tbl_info.removeRow(rowid)

    def upload(self, list): 
        # # 업로드 할 잔업시간 값이 float 형식이 아니면 중지
        # for i in list:
        #     try:
        #         float(i[7])
        #     except:
        #         self.msg_box("입력오류", "잔업시간 값이 숫자가 아닙니다.")
        #         return

        from db.db_insert import Insert
        data_insert = Insert()
        result = data_insert.insert_prod_info(list)

        self.msg_box(result[0], result[1])
        self.txt_select_file.setText("")
        #테이블을 지울경우 clear 같은 종류의 실행문을 쓰면 row 또는 column 헤더가 남는 경우가 생긴다.
        #column과 row 갯수를 "0"로 만들면 이런 현상이 생기지 않는다.
        self.tbl_info.setColumnCount(0)
        self.tbl_info.setRowCount(0)

        # # 부서명 가져오기 팝업
        # def popup_dept_info(self):
        #     input_dialog = InputWindow()
        #     if input_dialog.exec_():
        #         value = input_dialog.get_input_value()

        #     try:
        #         self.txt_dept_id.setText(value[0].text())
        #         self.txt_dept_name.setText(value[1].text())
        #     except:
        #         return
            
        # def dept_name(self, arg_1):  
        #     self.txt_dept_id.setText("arg_1.text()")
        #     print(arg_1)

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