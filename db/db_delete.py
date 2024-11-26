import sys
sys.path.append("./db") #모듈 import가 안 될때 경로를 지정해 주기
import db.db_info as conn_info
import pymysql
from PyQt5.QtWidgets import QMessageBox

class Delete:
    def __init__(self):
        db_info = conn_info.Connect()

        self.host = db_info.host
        self.user = db_info.username
        self.passwd = db_info.password
        self.db = db_info.database
        self.port = db_info.port

        self.conn = pymysql.connect(host=self.host, user=self.user, passwd=self.passwd, db=self.db, port=self.port, use_unicode=True, charset='utf8')        

    def delete_emp_overtime(self, arg_1):
        cursor = self.conn.cursor()

        try:
            # pymysql을 통해 쿼리 입력할 때, 아래와 같은 오류 문구를 만나곤 한다.
            # ValueError: unsupported format character 'Y' (0x59) at index
            # 이는 쿼리의 변수 표현에 쓰이는 %s와 data format 변경하는 (예시에서는 DATE_FORMAT) 에서의 %를 구분해주지 않았기 때문이다.
            # SELECT DATE_FORMAT(DeviceReportedTime, '%Y-%m-%d %H:%i:%s') AS date, Facility, Priority, FromHost, FromIP, Message FROM SystemEvents WHERE DeviceReportedTime BETWEEN '%s 00:00:00' AND '%s 23:59:59' ORDER BY DeviceReportedTime DESC
            # 위와 같이 작성하면 오류가 발생하는 것이다.            
            # DATE_FORMAT 안의 %를 %%로 변경해주어 아래와 같은 코드로 변경해주자. 

            query = """DELETE FROM overtime WHERE id = %s;""" #날짜를 비교 하기 위해 안쪽 select문 사용, qt 테이블 입력을 위해 날짜 형식을 문자로 바꾸려고 밖의 select문 사용
            cursor.execute(query, arg_1) #excute 문에 조회용 변수를 전달 할 때는 튜블 또는 리스트로 !!!!
            self.conn.commit()
            self.conn.close()

            self.msg_box("삭제결과", "정상적으로 삭제 되었습니다.")            

        except Exception as e:
            self.conn.close()
            self.msg_box("Error", str(e))

    def msg_box(self, msg_1, msg_2):
        msg = QMessageBox()
        msg.setWindowTitle(msg_1)               # 제목설정
        msg.setText(msg_2)                          # 내용설정
        msg.exec_()                                 # 메세지박스 실행

  