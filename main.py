import os
import sys
from PyQt5.QtWidgets import *
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
main_window= uic.loadUiType(resource_path("./ui/main_window.ui"))[0] # Window 사용시 ui 주소

#화면을 띄우는데 사용되는 Class 선언
class WindowClass(QMainWindow, main_window) :
    def __init__(self) :
        super().__init__()
        self.version = 1.0
        self.slots()

        self.mainwindow()

    def slots(self):
        pass

    def mainwindow(self):
       
        self.setupUi(self)
        self.setWindowTitle(f"DOOCHPUMP MES_v{self.version}")

        menu_bar = self.menuBar()

        # 메뉴바에 스타일시트 적용 (옅은 회색) 
        menu_bar.setStyleSheet("background-color: #F0F0F0; color: black;")

        # 메뉴바에 메뉴 추가
        prod_menu = menu_bar.addMenu("생산오더")
        overtime_menu = menu_bar.addMenu("잔업시간")

        # 생산관리
        prod_present = QAction('작업현황', self)
        prod_present.setStatusTip("작업현황")
        prod_present.triggered.connect(self.prod_present)

        prod_stop_present = QAction('중지사유', self)
        prod_stop_present.setStatusTip("중지사유")
        prod_stop_present.triggered.connect(self.prod_stop_present)

        prod_order_upload = QAction('생산오더_업로드', self)
        prod_order_upload.setStatusTip("생산오더_업로드")
        prod_order_upload.triggered.connect(self.prod_order_upload)

        # 잔업관리
        overtime_present = QAction('잔업현황', self)
        overtime_present.setStatusTip("잔업현황")
        overtime_present.triggered.connect(self.overtime_present)

        # select_dept = QAction('부서별 조회', self)
        # select_dept.setStatusTip("부서별 조회")
        # select_dept.triggered.connect(self.select_dept)

        # select_emp = QAction('사원별 조회', self)
        # select_emp.setStatusTip("사원별 조회")
        # select_emp.triggered.connect(self.select_emp)

        # select_month = QAction('월/사원별 조회', self)
        # select_month.setStatusTip("월/사원별 조회")
        # select_month.triggered.connect(self.select_month)

        # update_emp = QAction('잔업시간 수정', self)
        # update_emp.setStatusTip("잔업시간 수정")
        # update_emp.triggered.connect(self.update_emp)
        
        # input_emp = QAction('잔업시간 입력', self)
        # input_emp.setStatusTip("잔업시간 입력")
        # input_emp.triggered.connect(self.input_emp)

        # upload_overtime = QAction('잔업시간 업로드', self)
        # upload_overtime.setStatusTip("잔업시간 업로드")
        # upload_overtime.triggered.connect(self.upload_overtime)

        # emp_master = QAction('인사정보', self)
        # emp_master.setStatusTip("인사정보")
        # emp_master.triggered.connect(self.emp_master)

        prod_menu.addAction(prod_present)
        prod_menu.addAction(prod_stop_present)
        prod_menu.addSeparator()
        prod_menu.addAction(prod_order_upload)

        overtime_menu.addAction(overtime_present)

        status_bar = self.statusBar()
        self.setStatusBar(status_bar)

    def prod_present(self):
        import main_prod as main_prod_window

        self.total_prod_window = main_prod_window.MainWindow()
        self.setCentralWidget(self.total_prod_window)
        self.show()

    def prod_stop_present(self):
        import stop_prod as stop_prod_window

        self.stop_prod_window = stop_prod_window.MainWindow()
        self.setCentralWidget(self.stop_prod_window)
        self.show()

    def prod_order_upload(self):
        import upload_prod as upload_prod_window

        self.upload_prod_window = upload_prod_window.MainWindow()
        self.setCentralWidget(self.upload_prod_window)
        self.show()

    def overtime_present(self):
        import main_overtime as main_overtime_window

        self.main_overtime_window = main_overtime_window.MainWindow()
        self.setCentralWidget(self.main_overtime_window)
        self.show()

    def select_dept(self):
        import dept_overtime as select_dept_window

        self.dept_window = select_dept_window.DeptMainWindow()
        self.dept_window.show()
    
    def select_emp(self):
        import emp_overtime as select_emp_window

        self.emp_window = select_emp_window.MainWindow()
        self.emp_window.show() 

    def select_month(self):
        import main_overtime as select_emp_month_window

        self.emp_month_window = select_emp_month_window.MainWindow()
        self.emp_month_window.show()     
    
    def update_emp(self):
        import emp_overtime_update as update_emp_window

        self.emp_update_window = update_emp_window.MainWindow()
        self.emp_update_window.show() 

    def input_emp(self):
        import emp_overtime_input as input_emp_window

        self.emp_input_window = input_emp_window.MainWindow()
        self.emp_input_window.show() 

    def upload_overtime(self):
        import upload as upload_window

        self.upload_window = upload_window.MainWindow()
        self.upload_window.show()

    def emp_master(self):
        import emp_info as emp_info

        self.emp_master = emp_info.MainWindow()
        self.emp_master.show()

    def window_close(self):
        self.close()

 
    def msg_box(self, arg_1, arg_2):
        msg = QMessageBox()
        msg.setWindowTitle(arg_1)               
        msg.setText(arg_2)                          
        msg.exec_()       

if __name__ == "__main__" :
    app = QApplication(sys.argv)
    try:
        myWindow = WindowClass()
        myWindow.show()
        app.exec_()
    except Exception as e:
        msg = QMessageBox()
        msg.setWindowTitle("Error")               
        msg.setText(str(e))                          
        msg.exec_()  