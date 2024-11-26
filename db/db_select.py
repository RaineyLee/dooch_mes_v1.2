import sys
sys.path.append("./db") #모듈 import가 안 될때 경로를 지정해 주기
import db.db_info as conn_info
import pymysql
from PyQt5.QtWidgets import QMessageBox

class Select:

    def __init__(self):
        db_info = conn_info.Connect()

        self.host = db_info.host
        self.user = db_info.username
        self.passwd = db_info.password
        self.db = db_info.database
        self.port = db_info.port

        self.conn = pymysql.connect(host=self.host, user=self.user, passwd=self.passwd, db=self.db, port=self.port, use_unicode=True, charset='utf8')  

    def select_version(self):
        cursor = self.conn.cursor()

        try:
            # pymysql을 통해 쿼리 입력할 때, 아래와 같은 오류 문구를 만나곤 한다.
            # ValueError: unsupported format character 'Y' (0x59) at index
            # 이는 쿼리의 변수 표현에 쓰이는 %s와 data format 변경하는 (예시에서는 DATE_FORMAT) 에서의 %를 구분해주지 않았기 때문이다.
            # SELECT DATE_FORMAT(DeviceReportedTime, '%Y-%m-%d %H:%i:%s') AS date, Facility, Priority, FromHost, FromIP, Message FROM SystemEvents WHERE DeviceReportedTime BETWEEN '%s 00:00:00' AND '%s 23:59:59' ORDER BY DeviceReportedTime DESC
            # 위와 같이 작성하면 오류가 발생하는 것이다.            
            # DATE_FORMAT 안의 %를 %%로 변경해주어 아래와 같은 코드로 변경해주자. 

            query = """SELECT round(max(no),2) FROM version;;""" #날짜를 비교 하기 위해 안쪽 select문 사용, qt 테이블 입력을 위해 날짜 형식을 문자로 바꾸려고 밖의 select문 사용
            cursor.execute(query) #excute 문에 조회용 변수를 전달 할 때는 튜블 또는 리스트로 !!!!
            result = cursor.fetchall()
            return result

        except Exception as e:
            self.msg_box("Error", str(e))

    def select_location(self, arr_1):
        cursor = self.conn.cursor()

        try:
            # pymysql을 통해 쿼리 입력할 때, 아래와 같은 오류 문구를 만나곤 한다.
            # ValueError: unsupported format character 'Y' (0x59) at index
            # 이는 쿼리의 변수 표현에 쓰이는 %s와 data format 변경하는 (예시에서는 DATE_FORMAT) 에서의 %를 구분해주지 않았기 때문이다.
            # SELECT DATE_FORMAT(DeviceReportedTime, '%Y-%m-%d %H:%i:%s') AS date, Facility, Priority, FromHost, FromIP, Message FROM SystemEvents WHERE DeviceReportedTime BETWEEN '%s 00:00:00' AND '%s 23:59:59' ORDER BY DeviceReportedTime DESC
            # 위와 같이 작성하면 오류가 발생하는 것이다.            
            # DATE_FORMAT 안의 %를 %%로 변경해주어 아래와 같은 코드로 변경해주자. 

            query = """SELECT id, name, location, brand, package, DATE_FORMAT(c_date, '%%Y-%%m-%%d'), DATE_FORMAT(u_date, '%%Y-%%m-%%d') FROM item_location WHERE u_date BETWEEN %s AND %s;""" #날짜를 비교 하기 위해 안쪽 select문 사용, qt 테이블 입력을 위해 날짜 형식을 문자로 바꾸려고 밖의 select문 사용
            cursor.execute(query, arr_1) #excute 문에 조회용 변수를 전달 할 때는 튜블 또는 리스트로 !!!!
            result = cursor.fetchall()

            if result:
                # self.conn.close()
                # self.msg_box("조회완료", "정상적으로 조회 되었습니다.")
                return result
            else:
                self.conn.close()
                self.msg_box("조회결과", "조회결과가 없습니다.")
                return            

        except Exception as e:
            self.msg_box("Error", str(e))

        return result

    def select_department(self):
        cursor = self.conn.cursor()

        try:
            # pymysql을 통해 쿼리 입력할 때, 아래와 같은 오류 문구를 만나곤 한다.
            # ValueError: unsupported format character 'Y' (0x59) at index
            # 이는 쿼리의 변수 표현에 쓰이는 %s와 data format 변경하는 (예시에서는 DATE_FORMAT) 에서의 %를 구분해주지 않았기 때문이다.
            # SELECT DATE_FORMAT(DeviceReportedTime, '%Y-%m-%d %H:%i:%s') AS date, Facility, Priority, FromHost, FromIP, Message FROM SystemEvents WHERE DeviceReportedTime BETWEEN '%s 00:00:00' AND '%s 23:59:59' ORDER BY DeviceReportedTime DESC
            # 위와 같이 작성하면 오류가 발생하는 것이다.            
            # DATE_FORMAT 안의 %를 %%로 변경해주어 아래와 같은 코드로 변경해주자. 

            query = """SELECT dept_id, dept_name FROM department WHERE yn = "y" ORDER BY dept_id;""" #날짜를 비교 하기 위해 안쪽 select문 사용, qt 테이블 입력을 위해 날짜 형식을 문자로 바꾸려고 밖의 select문 사용
            cursor.execute(query) #excute 문에 조회용 변수를 전달 할 때는 튜블 또는 리스트로 !!!!
            result = cursor.fetchall()

            if result:
                self.conn.close()
                return result
            else:
                self.conn.close()
                return            

        except Exception as e:
            self.msg_box("Error", str(e))
    
    def select_employee(self, arg_1):
        cursor = self.conn.cursor()

        try:
            # pymysql을 통해 쿼리 입력할 때, 아래와 같은 오류 문구를 만나곤 한다.
            # ValueError: unsupported format character 'Y' (0x59) at index
            # 이는 쿼리의 변수 표현에 쓰이는 %s와 data format 변경하는 (예시에서는 DATE_FORMAT) 에서의 %를 구분해주지 않았기 때문이다.
            # SELECT DATE_FORMAT(DeviceReportedTime, '%Y-%m-%d %H:%i:%s') AS date, Facility, Priority, FromHost, FromIP, Message FROM SystemEvents WHERE DeviceReportedTime BETWEEN '%s 00:00:00' AND '%s 23:59:59' ORDER BY DeviceReportedTime DESC
            # 위와 같이 작성하면 오류가 발생하는 것이다.            
            # DATE_FORMAT 안의 %를 %%로 변경해주어 아래와 같은 코드로 변경해주자. 

            query = """SELECT a.dept_id, a.dept_name, b.emp_id, b.emp_name 
                        FROM department a, employee b 
                        WHERE b.dept_id = a.dept_id 
                        AND a.dept_id = %s
                        ORDER BY b.emp_id;""" #날짜를 비교 하기 위해 안쪽 select문 사용, qt 테이블 입력을 위해 날짜 형식을 문자로 바꾸려고 밖의 select문 사용
            cursor.execute(query, arg_1) #excute 문에 조회용 변수를 전달 할 때는 튜블 또는 리스트로 !!!!
            result = cursor.fetchall()

            if result:
                self.conn.close()
                return result
            else:
                self.conn.close()
            

        except Exception as e:
            self.msg_box("Error", str(e))

    def dept_overtime_1(self, arr_1):
        cursor = self.conn.cursor()

        try:
            # pymysql을 통해 쿼리 입력할 때, 아래와 같은 오류 문구를 만나곤 한다.
            # ValueError: unsupported format character 'Y' (0x59) at index
            # 이는 쿼리의 변수 표현에 쓰이는 %s와 data format 변경하는 (예시에서는 DATE_FORMAT) 에서의 %를 구분해주지 않았기 때문이다.
            # SELECT DATE_FORMAT(DeviceReportedTime, '%Y-%m-%d %H:%i:%s') AS date, Facility, Priority, FromHost, FromIP, Message FROM SystemEvents WHERE DeviceReportedTime BETWEEN '%s 00:00:00' AND '%s 23:59:59' ORDER BY DeviceReportedTime DESC
            # 위와 같이 작성하면 오류가 발생하는 것이다.            
            # DATE_FORMAT 안의 %를 %%로 변경해주어 아래와 같은 코드로 변경해주자. 
            #날짜를 비교 하기 위해 안쪽 select문 사용, qt 테이블 입력을 위해 날짜 형식을 문자로 바꾸려고 밖의 select문 사용

            query = """SELECT IFNULL(b.dept_name, "생산본부") as "Dept", IFNULL(a.yyyy_mm, "합계") AS "Month", round(SUM(a.overtime),2) AS "OVERTIME"
                        FROM overtime a, department b, employee c   
                        WHERE a.yyyy_mm BETWEEN %s AND %s
                        AND a.dept_id = b.dept_id
                        AND a.emp_id = c.emp_id
                        AND b.dept_id LIKE %s
                        GROUP BY b.dept_name, a.yyyy_mm
                        WITH ROLLUP;
                    """ 
            cursor.execute(query, arr_1) #excute 문에 조회용 변수를 전달 할 때는 튜블 또는 리스트로 !!!!
            result = cursor.fetchall()

            if result:
                self.conn.close()
                return result
            else:
                self.conn.close()
                self.msg_box("조회결과", "조회결과가 없습니다.")
                return            

        except Exception as e:
            self.msg_box("Error", str(e))

    def dept_overtime_2(self, arr_1):
        cursor = self.conn.cursor()

        try:
            # pymysql을 통해 쿼리 입력할 때, 아래와 같은 오류 문구를 만나곤 한다.
            # ValueError: unsupported format character 'Y' (0x59) at index
            # 이는 쿼리의 변수 표현에 쓰이는 %s와 data format 변경하는 (예시에서는 DATE_FORMAT) 에서의 %를 구분해주지 않았기 때문이다.
            # SELECT DATE_FORMAT(DeviceReportedTime, '%Y-%m-%d %H:%i:%s') AS date, Facility, Priority, FromHost, FromIP, Message FROM SystemEvents WHERE DeviceReportedTime BETWEEN '%s 00:00:00' AND '%s 23:59:59' ORDER BY DeviceReportedTime DESC
            # 위와 같이 작성하면 오류가 발생하는 것이다.            
            # DATE_FORMAT 안의 %를 %%로 변경해주어 아래와 같은 코드로 변경해주자. 
            #날짜를 비교 하기 위해 안쪽 select문 사용, qt 테이블 입력을 위해 날짜 형식을 문자로 바꾸려고 밖의 select문 사용

            query = """SELECT IFNULL(b.dept_name, "") as "Dept", IFNULL(a.yyyy_mm, "합계") AS "Month", round(SUM(a.overtime),2) AS "OVERTIME"
                        FROM overtime a, department b, employee c   
                        WHERE a.yyyy_mm BETWEEN %s AND %s
                        AND a.dept_id = b.dept_id
                        AND a.emp_id = c.emp_id
                        AND b.dept_id LIKE %s
                        GROUP BY a.yyyy_mm
                        WITH ROLLUP;
                    """ 
            cursor.execute(query, arr_1) #excute 문에 조회용 변수를 전달 할 때는 튜블 또는 리스트로 !!!!
            result = cursor.fetchall()

            if result:
                self.conn.close()
                return result
            else:
                self.conn.close()
                self.msg_box("조회결과", "조회결과가 없습니다.")
                return            

        except Exception as e:
            self.msg_box("Error", str(e))
    
    def emp_overtime_1(self, arr_1):
        cursor = self.conn.cursor()

        try:
            # pymysql을 통해 쿼리 입력할 때, 아래와 같은 오류 문구를 만나곤 한다.
            # ValueError: unsupported format character 'Y' (0x59) at index
            # 이는 쿼리의 변수 표현에 쓰이는 %s와 data format 변경하는 (예시에서는 DATE_FORMAT) 에서의 %를 구분해주지 않았기 때문이다.
            # SELECT DATE_FORMAT(DeviceReportedTime, '%Y-%m-%d %H:%i:%s') AS date, Facility, Priority, FromHost, FromIP, Message FROM SystemEvents WHERE DeviceReportedTime BETWEEN '%s 00:00:00' AND '%s 23:59:59' ORDER BY DeviceReportedTime DESC
            # 위와 같이 작성하면 오류가 발생하는 것이다.            
            # DATE_FORMAT 안의 %를 %%로 변경해주어 아래와 같은 코드로 변경해주자. 
            #날짜를 비교 하기 위해 안쪽 select문 사용, qt 테이블 입력을 위해 날짜 형식을 문자로 바꾸려고 밖의 select문 사용

            query = """SELECT ifnull(b.dept_name, "생산본부") as "Dept",  IFNULL(c.emp_name, "") AS "Name",IFNULL(a.yyyy_mm, "합계") AS "Month", round(SUM(a.overtime),2) AS "OVERTIME"
                        FROM overtime a, department b, employee c   
                        WHERE a.yyyy_mm BETWEEN %s AND %s
                        AND a.dept_id = b.dept_id
                        AND a.emp_id = c.emp_id
                        AND a.dept_id = %s
                        GROUP BY c.emp_name, a.yyyy_mm
                        WITH ROLLUP;""" 
            cursor.execute(query, arr_1) #excute 문에 조회용 변수를 전달 할 때는 튜블 또는 리스트로 !!!!
            result = cursor.fetchall()

            if result:
                self.conn.close()
                return result
            else:
                self.conn.close()
                self.msg_box("조회결과", "조회결과가 없습니다.")
                return            

        except Exception as e:
            self.msg_box("Error", str(e))

    def emp_overtime_2(self, arr_1):
        cursor = self.conn.cursor()

        try:
            # pymysql을 통해 쿼리 입력할 때, 아래와 같은 오류 문구를 만나곤 한다.
            # ValueError: unsupported format character 'Y' (0x59) at index
            # 이는 쿼리의 변수 표현에 쓰이는 %s와 data format 변경하는 (예시에서는 DATE_FORMAT) 에서의 %를 구분해주지 않았기 때문이다.
            # SELECT DATE_FORMAT(DeviceReportedTime, '%Y-%m-%d %H:%i:%s') AS date, Facility, Priority, FromHost, FromIP, Message FROM SystemEvents WHERE DeviceReportedTime BETWEEN '%s 00:00:00' AND '%s 23:59:59' ORDER BY DeviceReportedTime DESC
            # 위와 같이 작성하면 오류가 발생하는 것이다.            
            # DATE_FORMAT 안의 %를 %%로 변경해주어 아래와 같은 코드로 변경해주자. 
            #날짜를 비교 하기 위해 안쪽 select문 사용, qt 테이블 입력을 위해 날짜 형식을 문자로 바꾸려고 밖의 select문 사용

            query = """SELECT ifnull(b.dept_name, "생산본부") as "Dept",  IFNULL(c.emp_name, "") AS "Name",IFNULL(a.yyyy_mm, "합계") AS "Month", round(SUM(a.overtime),2) AS "OVERTIME"
                        FROM overtime a, department b, employee c   
                        WHERE a.yyyy_mm BETWEEN %s AND %s
                        AND a.dept_id = b.dept_id
                        AND a.emp_id = c.emp_id
                        AND c.emp_id = %s
                        GROUP BY a.yyyy_mm
                        WITH ROLLUP;""" 
            cursor.execute(query, arr_1) #excute 문에 조회용 변수를 전달 할 때는 튜블 또는 리스트로 !!!!
            result = cursor.fetchall()

            if result:
                self.conn.close()
                return result
            else:
                self.conn.close()
                self.msg_box("조회결과", "조회결과가 없습니다.")
                return            

        except Exception as e:
            self.msg_box("Error", str(e))

    def all_overtime_1(self, arr_1):
        cursor = self.conn.cursor()

        try:
            # pymysql을 통해 쿼리 입력할 때, 아래와 같은 오류 문구를 만나곤 한다.
            # ValueError: unsupported format character 'Y' (0x59) at index
            # 이는 쿼리의 변수 표현에 쓰이는 %s와 data format 변경하는 (예시에서는 DATE_FORMAT) 에서의 %를 구분해주지 않았기 때문이다.
            # SELECT DATE_FORMAT(DeviceReportedTime, '%Y-%m-%d %H:%i:%s') AS date, Facility, Priority, FromHost, FromIP, Message FROM SystemEvents WHERE DeviceReportedTime BETWEEN '%s 00:00:00' AND '%s 23:59:59' ORDER BY DeviceReportedTime DESC
            # 위와 같이 작성하면 오류가 발생하는 것이다.            
            # DATE_FORMAT 안의 %를 %%로 변경해주어 아래와 같은 코드로 변경해주자. 

            query = """
                    SELECT a.dept_id AS "DeptID", a.dept_name AS "DeptNAME", a.emp_id AS "EmpID", a.emp_name AS "Name", 
                    a.overtime_date AS "Month", a.s_time AS "START", a.t_time AS "END", round(a.overtime,2) AS "OVERTIME", a.detail AS "DETAIL", a.note AS "NOTE"
                    FROM overtime a  
                    WHERE a.dept_id = %s
                    AND a.overtime_date BETWEEN %s AND %s
                    ORDER BY  a.overtime_date, a.dept_id, a.emp_id
                    ;               
                    """ 
                    # 월로 비교하기 AND DATE_FORMAT(a.overtime_date, "%%Y-%%m") BETWEEN %s AND %s
                                #날짜를 비교 하기 위해 안쪽 select문 사용, qt 테이블 입력을 위해 날짜 형식을 문자로 바꾸려고 밖의 select문 사용
            cursor.execute(query, arr_1) #excute 문에 조회용 변수를 전달 할 때는 튜블 또는 리스트로 !!!!
            result = cursor.fetchall()

            if result:
                self.conn.close()
                return result
            else:
                self.conn.close()
                self.msg_box("조회결과", "조회결과가 없습니다.")
                return            

        except Exception as e:
            self.msg_box("Error", str(e))

    def all_overtime_2(self, arr_1):
        cursor = self.conn.cursor()

        try:
            # pymysql을 통해 쿼리 입력할 때, 아래와 같은 오류 문구를 만나곤 한다.
            # ValueError: unsupported format character 'Y' (0x59) at index
            # 이는 쿼리의 변수 표현에 쓰이는 %s와 data format 변경하는 (예시에서는 DATE_FORMAT) 에서의 %를 구분해주지 않았기 때문이다.
            # SELECT DATE_FORMAT(DeviceReportedTime, '%Y-%m-%d %H:%i:%s') AS date, Facility, Priority, FromHost, FromIP, Message FROM SystemEvents WHERE DeviceReportedTime BETWEEN '%s 00:00:00' AND '%s 23:59:59' ORDER BY DeviceReportedTime DESC
            # 위와 같이 작성하면 오류가 발생하는 것이다.            
            # DATE_FORMAT 안의 %를 %%로 변경해주어 아래와 같은 코드로 변경해주자. 

            query = """
                    SELECT b.dept_id AS "DeptID", b.dept_name AS "DeptNAME", c.emp_id AS "EmpID", c.emp_name AS "Name", 
                    a.overtime_date AS "Month", round(a.overtime,2) AS "OVERTIME", a.s_time AS "START", a.t_time AS "END", a.detail AS "DETAIL", a.note AS "NOTE"
                    FROM overtime a, department b, employee c   
                    WHERE a.emp_id = c.emp_id
                    AND a.dept_id = b.dept_id
                    AND a.overtime_date BETWEEN %s AND %s 
                    ORDER BY  a.overtime_date, b.dept_id, c.emp_id
                    ;               
                    """ 
                    # 월로 비교 하기 AND DATE_FORMAT(a.overtime_date, "%%Y-%%m") BETWEEN %s AND %s
                                #날짜를 비교 하기 위해 안쪽 select문 사용, qt 테이블 입력을 위해 날짜 형식을 문자로 바꾸려고 밖의 select문 사용
            cursor.execute(query, arr_1) #excute 문에 조회용 변수를 전달 할 때는 튜블 또는 리스트로 !!!!
            result = cursor.fetchall()

            if result:
                self.conn.close()
                return result
            else:
                self.conn.close()
                self.msg_box("조회결과", "조회결과가 없습니다.")
                return            

        except Exception as e:
            self.msg_box("Error", str(e))

    def emp_info(self, arg):
        cursor = self.conn.cursor()

            # pymysql을 통해 쿼리 입력할 때, 아래와 같은 오류 문구를 만나곤 한다.
            # ValueError: unsupported format character 'Y' (0x59) at index
            # 이는 쿼리의 변수 표현에 쓰이는 %s와 data format 변경하는 (예시에서는 DATE_FORMAT) 에서의 %를 구분해주지 않았기 때문이다.
            # SELECT DATE_FORMAT(DeviceReportedTime, '%Y-%m-%d %H:%i:%s') AS date, Facility, Priority, FromHost, FromIP, Message FROM SystemEvents WHERE DeviceReportedTime BETWEEN '%s 00:00:00' AND '%s 23:59:59' ORDER BY DeviceReportedTime DESC
            # 위와 같이 작성하면 오류가 발생하는 것이다.            
            # DATE_FORMAT 안의 %를 %%로 변경해주어 아래와 같은 코드로 변경해주자. 
        try:
            query = """
                    SELECT a.dept_id, b.dept_name, a.emp_id, a.emp_name, a.yn
                    FROM employee a, department b
                    WHERE a.dept_id = b.dept_id
                    AND a.yn = %s
                    ORDER BY a.dept_id, a.emp_id
                    ;               
                    """ 
                                #날짜를 비교 하기 위해 안쪽 select문 사용, qt 테이블 입력을 위해 날짜 형식을 문자로 바꾸려고 밖의 select문 사용
            cursor.execute(query, arg) #excute 문에 조회용 변수를 전달 할 때는 튜블 또는 리스트로 !!!!
            result = cursor.fetchall()

            if result:
                self.conn.close()
                return result
            else:
                self.conn.close()
                self.msg_box("조회결과", "조회결과가 없습니다.")
                return            

        except Exception as e:
            self.msg_box("Error", str(e))

    def emp_info_dept(self, arr_1):
        cursor = self.conn.cursor()

        try:
            # pymysql을 통해 쿼리 입력할 때, 아래와 같은 오류 문구를 만나곤 한다.
            # ValueError: unsupported format character 'Y' (0x59) at index
            # 이는 쿼리의 변수 표현에 쓰이는 %s와 data format 변경하는 (예시에서는 DATE_FORMAT) 에서의 %를 구분해주지 않았기 때문이다.
            # SELECT DATE_FORMAT(DeviceReportedTime, '%Y-%m-%d %H:%i:%s') AS date, Facility, Priority, FromHost, FromIP, Message FROM SystemEvents WHERE DeviceReportedTime BETWEEN '%s 00:00:00' AND '%s 23:59:59' ORDER BY DeviceReportedTime DESC
            # 위와 같이 작성하면 오류가 발생하는 것이다.            
            # DATE_FORMAT 안의 %를 %%로 변경해주어 아래와 같은 코드로 변경해주자. 

            query = """
                    SELECT a.dept_id, b.dept_name, a.emp_id, a.emp_name, a.yn
                    FROM employee a, department b
                    WHERE a.dept_id = b.dept_id AND b.dept_id = %s AND a.yn = %s
                    ORDER BY a.dept_id, a.emp_id
                    ;               
                    """ 
                                #날짜를 비교 하기 위해 안쪽 select문 사용, qt 테이블 입력을 위해 날짜 형식을 문자로 바꾸려고 밖의 select문 사용
            cursor.execute(query, arr_1) #excute 문에 조회용 변수를 전달 할 때는 튜블 또는 리스트로 !!!!
            result = cursor.fetchall()

            if result:
                self.conn.close()
                return result
            else:
                self.conn.close()
                self.msg_box("조회결과", "조회결과가 없습니다.")
                return            

        except Exception as e:
            self.msg_box("Error", str(e))

    def emp_info_dept_emp(self, arr_1):
        cursor = self.conn.cursor()

        try:
            # pymysql을 통해 쿼리 입력할 때, 아래와 같은 오류 문구를 만나곤 한다.
            # ValueError: unsupported format character 'Y' (0x59) at index
            # 이는 쿼리의 변수 표현에 쓰이는 %s와 data format 변경하는 (예시에서는 DATE_FORMAT) 에서의 %를 구분해주지 않았기 때문이다.
            # SELECT DATE_FORMAT(DeviceReportedTime, '%Y-%m-%d %H:%i:%s') AS date, Facility, Priority, FromHost, FromIP, Message FROM SystemEvents WHERE DeviceReportedTime BETWEEN '%s 00:00:00' AND '%s 23:59:59' ORDER BY DeviceReportedTime DESC
            # 위와 같이 작성하면 오류가 발생하는 것이다.            
            # DATE_FORMAT 안의 %를 %%로 변경해주어 아래와 같은 코드로 변경해주자. 

            query = """
                    SELECT a.dept_id, b.dept_name, a.emp_id, a.emp_name, a.yn
                    FROM employee a, department b
                    WHERE a.dept_id = b.dept_id AND a.emp_id = %s AND a.yn = %s
                    ORDER BY a.dept_id, a.emp_id
                    ;               
                    """ 
                                #날짜를 비교 하기 위해 안쪽 select문 사용, qt 테이블 입력을 위해 날짜 형식을 문자로 바꾸려고 밖의 select문 사용
            cursor.execute(query, arr_1) #excute 문에 조회용 변수를 전달 할 때는 튜블 또는 리스트로 !!!!
            result = cursor.fetchall()

            if result:
                self.conn.close()
                return result
            else:
                self.conn.close()
                self.msg_box("조회결과", "조회결과가 없습니다.")
                return            

        except Exception as e:
            self.msg_box("Error", str(e))

    def update_overtime(self, arr_1):
        cursor = self.conn.cursor()

        try:
            query = """
                    SELECT t.id, t.dept_id, t.dept_name, t.emp_id, t.emp_name, t.overtime_date, t.overtime, t.start, t.end, t.detail, t.note
                    FROM(
                    SELECT a.id AS "id", b.dept_id AS "dept_id", b.dept_name AS "dept_name", c.emp_id AS "emp_id", c.emp_name AS "emp_name", 
                    a.overtime_date AS "overtime_date", round(a.overtime,2) AS "overtime", a.s_time AS "start", a.t_time AS "end", a.detail AS "detail", a.note AS "note"
                    FROM overtime a, department b, employee c   
                    WHERE a.emp_id = c.emp_id
                    AND a.dept_id = b.dept_id
                    AND a.overtime_date BETWEEN %s AND %s 
                    ORDER BY  b.dept_id, c.emp_id, a.overtime_date) t
                    WHERE t.dept_id LIKE %s
                    AND t.emp_id LIKE %s
                    ;                   
                    """ 
                    # 월로 비교 하기 AND DATE_FORMAT(a.overtime_date, "%%Y-%%m") BETWEEN %s AND %s
                                #날짜를 비교 하기 위해 안쪽 select문 사용, qt 테이블 입력을 위해 날짜 형식을 문자로 바꾸려고 밖의 select문 사용
            cursor.execute(query, arr_1) #excute 문에 조회용 변수를 전달 할 때는 튜블 또는 리스트로 !!!!
            result = cursor.fetchall()

            if result:
                self.conn.close()
                return result
            else:
                self.conn.close()
                self.msg_box("조회결과", "조회결과가 없습니다.")
                return 
        except Exception as e:
            self.msg_box("Error", str(e))

    def update_overtime_id(self, arg_1):
        cursor = self.conn.cursor()

        try:
            # pymysql을 통해 쿼리 입력할 때, 아래와 같은 오류 문구를 만나곤 한다.
            # ValueError: unsupported format character 'Y' (0x59) at index
            # 이는 쿼리의 변수 표현에 쓰이는 %s와 data format 변경하는 (예시에서는 DATE_FORMAT) 에서의 %를 구분해주지 않았기 때문이다.
            # SELECT DATE_FORMAT(DeviceReportedTime, '%Y-%m-%d %H:%i:%s') AS date, Facility, Priority, FromHost, FromIP, Message FROM SystemEvents WHERE DeviceReportedTime BETWEEN '%s 00:00:00' AND '%s 23:59:59' ORDER BY DeviceReportedTime DESC
            # 위와 같이 작성하면 오류가 발생하는 것이다.            
            # DATE_FORMAT 안의 %를 %%로 변경해주어 아래와 같은 코드로 변경해주자. 

            query = """
                    SELECT t.id, t.dept_id, t.dept_name, t.emp_id, t.emp_name, t.overtime_date, t.overtime, t.start, t.end, t.detail, t.note
                    FROM(
                    SELECT a.id AS "id", b.dept_id AS "dept_id", b.dept_name AS "dept_name", c.emp_id AS "emp_id", c.emp_name AS "emp_name", 
                    a.overtime_date AS "overtime_date", round(a.overtime,2) AS "overtime", a.s_time AS "start", a.t_time AS "end", a.detail AS "detail", a.note AS "note"
                    FROM overtime a, department b, employee c   
                    WHERE a.emp_id = c.emp_id
                    AND a.dept_id = b.dept_id
                    ORDER BY  b.dept_id, c.emp_id, a.overtime_date) t
                    WHERE t.id = %s;
                    """ 
                    # 월로 비교 하기 AND DATE_FORMAT(a.overtime_date, "%%Y-%%m") BETWEEN %s AND %s
                                #날짜를 비교 하기 위해 안쪽 select문 사용, qt 테이블 입력을 위해 날짜 형식을 문자로 바꾸려고 밖의 select문 사용
            cursor.execute(query, arg_1) #excute 문에 조회용 변수를 전달 할 때는 튜블 또는 리스트로 !!!!
            result = cursor.fetchall()

            if result:
                self.conn.close()
                return result
            else:
                self.conn.close()
                self.msg_box("조회결과", "조회결과가 없습니다.")
                return            

        except Exception as e:
            self.msg_box("Error", str(e))

    def select_password(self, arg_1):
        cursor = self.conn.cursor()

        try:
            # pymysql을 통해 쿼리 입력할 때, 아래와 같은 오류 문구를 만나곤 한다.
            # ValueError: unsupported format character 'Y' (0x59) at index
            # 이는 쿼리의 변수 표현에 쓰이는 %s와 data format 변경하는 (예시에서는 DATE_FORMAT) 에서의 %를 구분해주지 않았기 때문이다.
            # SELECT DATE_FORMAT(DeviceReportedTime, '%Y-%m-%d %H:%i:%s') AS date, Facility, Priority, FromHost, FromIP, Message FROM SystemEvents WHERE DeviceReportedTime BETWEEN '%s 00:00:00' AND '%s 23:59:59' ORDER BY DeviceReportedTime DESC
            # 위와 같이 작성하면 오류가 발생하는 것이다.            
            # DATE_FORMAT 안의 %를 %%로 변경해주어 아래와 같은 코드로 변경해주자. 

            query = """
                    SELECT password
                    FROM login
                    WHERE id = %s;
                    """ 
                    # 월로 비교 하기 AND DATE_FORMAT(a.overtime_date, "%%Y-%%m") BETWEEN %s AND %s
                                #날짜를 비교 하기 위해 안쪽 select문 사용, qt 테이블 입력을 위해 날짜 형식을 문자로 바꾸려고 밖의 select문 사용
            cursor.execute(query, arg_1) #excute 문에 조회용 변수를 전달 할 때는 튜블 또는 리스트로 !!!!
            result = cursor.fetchone()

            if result:
                self.conn.close()
                return result
            else:
                self.conn.close()
                self.msg_box("조회결과", "조회결과가 없습니다.")
                return            

        except Exception as e:
            self.msg_box("Error", str(e))

    def select_dept_monthly(self):
        cursor = self.conn.cursor()

        try:
            # pymysql을 통해 쿼리 입력할 때, 아래와 같은 오류 문구를 만나곤 한다.
            # ValueError: unsupported format character 'Y' (0x59) at index
            # 이는 쿼리의 변수 표현에 쓰이는 %s와 data format 변경하는 (예시에서는 DATE_FORMAT) 에서의 %를 구분해주지 않았기 때문이다.
            # SELECT DATE_FORMAT(DeviceReportedTime, '%Y-%m-%d %H:%i:%s') AS date, Facility, Priority, FromHost, FromIP, Message FROM SystemEvents WHERE DeviceReportedTime BETWEEN '%s 00:00:00' AND '%s 23:59:59' ORDER BY DeviceReportedTime DESC
            # 위와 같이 작성하면 오류가 발생하는 것이다.            
            # DATE_FORMAT 안의 %를 %%로 변경해주어 아래와 같은 코드로 변경해주자. 

            query = """SELECT
                        dept_name AS "부서명",
                        SUM(CASE WHEN yyyy_mm = '2024-01' THEN overtime ELSE 0 END) AS '2024-01',
                        SUM(CASE WHEN yyyy_mm = '2024-02' THEN overtime ELSE 0 END) AS '2024-02',
                        SUM(CASE WHEN yyyy_mm = '2024-03' THEN overtime ELSE 0 END) AS '2024-03',
                        SUM(CASE WHEN yyyy_mm = '2024-04' THEN overtime ELSE 0 END) AS '2024-04',
                        SUM(CASE WHEN yyyy_mm = '2024-05' THEN overtime ELSE 0 END) AS '2024-05',
                        SUM(CASE WHEN yyyy_mm = '2024-06' THEN overtime ELSE 0 END) AS '2024-06',
                        SUM(CASE WHEN yyyy_mm = '2024-07' THEN overtime ELSE 0 END) AS '2024-07',
                        SUM(CASE WHEN yyyy_mm = '2024-08' THEN overtime ELSE 0 END) AS '2024-08',
                        SUM(CASE WHEN yyyy_mm = '2024-09' THEN overtime ELSE 0 END) AS '2024-09',
                        SUM(CASE WHEN yyyy_mm = '2024-10' THEN overtime ELSE 0 END) AS '2024-10',
                        SUM(CASE WHEN yyyy_mm = '2024-11' THEN overtime ELSE 0 END) AS '2024-11',
                        SUM(CASE WHEN yyyy_mm = '2024-12' THEN overtime ELSE 0 END) AS '2024-12'
                    FROM (    
                            SELECT * FROM overtime
                        ) T
                    GROUP BY dept_name
                    ORDER BY dept_name;""" 
                    # 월로 비교 하기 AND DATE_FORMAT(a.overtime_date, "%%Y-%%m") BETWEEN %s AND %s
                                #날짜를 비교 하기 위해 안쪽 select문 사용, qt 테이블 입력을 위해 날짜 형식을 문자로 바꾸려고 밖의 select문 사용
            cursor.execute(query) #excute 문에 조회용 변수를 전달 할 때는 튜블 또는 리스트로 !!!!
            result = cursor.fetchall()

            column_names = [description[0] for description in cursor.description] # db 컬럼명을 조회후 리스트로 만듬

            if result:
                self.conn.close()
                return result, column_names
            else:
                self.conn.close()
                self.msg_box("조회결과", "조회결과가 없습니다.")
                return            

        except Exception as e:
            self.msg_box("Error", str(e))

    def select_emp_monthly(self):
        cursor = self.conn.cursor()

        try:
            # pymysql을 통해 쿼리 입력할 때, 아래와 같은 오류 문구를 만나곤 한다.
            # ValueError: unsupported format character 'Y' (0x59) at index
            # 이는 쿼리의 변수 표현에 쓰이는 %s와 data format 변경하는 (예시에서는 DATE_FORMAT) 에서의 %를 구분해주지 않았기 때문이다.
            # SELECT DATE_FORMAT(DeviceReportedTime, '%Y-%m-%d %H:%i:%s') AS date, Facility, Priority, FromHost, FromIP, Message FROM SystemEvents WHERE DeviceReportedTime BETWEEN '%s 00:00:00' AND '%s 23:59:59' ORDER BY DeviceReportedTime DESC
            # 위와 같이 작성하면 오류가 발생하는 것이다.            
            # DATE_FORMAT 안의 %를 %%로 변경해주어 아래와 같은 코드로 변경해주자. 

            query = """SELECT
                        dept_name AS "부서명", emp_name AS "사원명",
                        SUM(CASE WHEN yyyy_mm = '2024-01' THEN overtime ELSE 0 END) AS '2024-01',
                        SUM(CASE WHEN yyyy_mm = '2024-02' THEN overtime ELSE 0 END) AS '2024-02',
                        SUM(CASE WHEN yyyy_mm = '2024-03' THEN overtime ELSE 0 END) AS '2024-03',
                        SUM(CASE WHEN yyyy_mm = '2024-04' THEN overtime ELSE 0 END) AS '2024-04',
                        SUM(CASE WHEN yyyy_mm = '2024-05' THEN overtime ELSE 0 END) AS '2024-05',
                        SUM(CASE WHEN yyyy_mm = '2024-06' THEN overtime ELSE 0 END) AS '2024-06',
                        SUM(CASE WHEN yyyy_mm = '2024-07' THEN overtime ELSE 0 END) AS '2024-07',
                        SUM(CASE WHEN yyyy_mm = '2024-08' THEN overtime ELSE 0 END) AS '2024-08',
                        SUM(CASE WHEN yyyy_mm = '2024-09' THEN overtime ELSE 0 END) AS '2024-09',
                        SUM(CASE WHEN yyyy_mm = '2024-10' THEN overtime ELSE 0 END) AS '2024-10',
                        SUM(CASE WHEN yyyy_mm = '2024-11' THEN overtime ELSE 0 END) AS '2024-11',
                        SUM(CASE WHEN yyyy_mm = '2024-12' THEN overtime ELSE 0 END) AS '2024-12'
                    FROM (    
                            SELECT * FROM overtime
                        ) T
                    GROUP BY dept_name, emp_name;""" 
            cursor.execute(query) #excute 문에 조회용 변수를 전달 할 때는 튜블 또는 리스트로 !!!!
            result = cursor.fetchall()

            column_names = [description[0] for description in cursor.description] # 컬럼명 조회후 리스트로 만듬

            if result:
                self.conn.close()
                return result, column_names
            else:
                self.conn.close()
                self.msg_box("조회결과", "조회결과가 없습니다.")
                return            

        except Exception as e:
            self.msg_box("Error", str(e))

    def select_monthly_sum(self):
        cursor = self.conn.cursor()

        try:
            # pymysql을 통해 쿼리 입력할 때, 아래와 같은 오류 문구를 만나곤 한다.
            # ValueError: unsupported format character 'Y' (0x59) at index
            # 이는 쿼리의 변수 표현에 쓰이는 %s와 data format 변경하는 (예시에서는 DATE_FORMAT) 에서의 %를 구분해주지 않았기 때문이다.
            # SELECT DATE_FORMAT(DeviceReportedTime, '%Y-%m-%d %H:%i:%s') AS date, Facility, Priority, FromHost, FromIP, Message FROM SystemEvents WHERE DeviceReportedTime BETWEEN '%s 00:00:00' AND '%s 23:59:59' ORDER BY DeviceReportedTime DESC
            # 위와 같이 작성하면 오류가 발생하는 것이다.            
            # DATE_FORMAT 안의 %를 %%로 변경해주어 아래와 같은 코드로 변경해주자. 

            query = """SELECT
                            "합계" AS "날짜",
                            SUM(CASE WHEN yyyy_mm = '2024-01' THEN overtime ELSE 0 END) AS '2024-01',
                            SUM(CASE WHEN yyyy_mm = '2024-02' THEN overtime ELSE 0 END) AS '2024-02',
                            SUM(CASE WHEN yyyy_mm = '2024-03' THEN overtime ELSE 0 END) AS '2024-03',
                            SUM(CASE WHEN yyyy_mm = '2024-04' THEN overtime ELSE 0 END) AS '2024-04',
                            SUM(CASE WHEN yyyy_mm = '2024-05' THEN overtime ELSE 0 END) AS '2024-05',
                            SUM(CASE WHEN yyyy_mm = '2024-06' THEN overtime ELSE 0 END) AS '2024-06',
                            SUM(CASE WHEN yyyy_mm = '2024-07' THEN overtime ELSE 0 END) AS '2024-07',
                            SUM(CASE WHEN yyyy_mm = '2024-08' THEN overtime ELSE 0 END) AS '2024-08',
                            SUM(CASE WHEN yyyy_mm = '2024-09' THEN overtime ELSE 0 END) AS '2024-09',
                            SUM(CASE WHEN yyyy_mm = '2024-10' THEN overtime ELSE 0 END) AS '2024-10',
                            SUM(CASE WHEN yyyy_mm = '2024-11' THEN overtime ELSE 0 END) AS '2024-11',
                            SUM(CASE WHEN yyyy_mm = '2024-12' THEN overtime ELSE 0 END) AS '2024-12'
                        FROM (
                            SELECT 
                                yyyy_mm,
                                SUM(overtime) AS overtime
                            FROM overtime
                            WHERE yyyy_mm BETWEEN '2024-01' AND '2024-12'  -- Ensure data is filtered for the year 2024
                            GROUP BY yyyy_mm
                        ) AS T
                        GROUP BY "합계";"""
            
            cursor.execute(query) #excute 문에 조회용 변수를 전달 할 때는 튜블 또는 리스트로 !!!!
            result = cursor.fetchall()

            column_names = [description[0] for description in cursor.description] # 컬럼명 조회후 리스트로 만듬

            if result:
                self.conn.close()
                return result, column_names
            else:
                self.conn.close()
                self.msg_box("조회결과", "조회결과가 없습니다.")
                return            

        except Exception as e:
            self.msg_box("Error", str(e))

    def msg_box(self, msg_1, msg_2):
        msg = QMessageBox()
        msg.setWindowTitle(msg_1)               # 제목설정
        msg.setText(msg_2)                          # 내용설정
        msg.exec_()     