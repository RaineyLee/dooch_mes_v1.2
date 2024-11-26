import sys
sys.path.append("./db") #모듈 import가 안 될때 경로를 지정해 주기
import db_info as conn_info
import pymysql

class Insert:

    def __init__(self):
        db_info = conn_info.Connect()

        self.host = db_info.host
        self.user = db_info.username
        self.passwd = db_info.password
        self.db = db_info.database
        self.port = db_info.port

        self.conn = pymysql.connect(host=self.host, user=self.user, passwd=self.passwd, db=self.db, port=self.port, use_unicode=True, charset='utf8')        

    # def insert_location(self, arr_1):
    #     cursor = self.conn.cursor()

    #     try:
    #         query = "TRUNCATE TABLE item_location;"
    #         cursor.execute(query)
    #         self.conn.commit()

    #         query = """INSERT INTO item_location (id, name, location, brand, package, c_date) VALUES (%s, %s, %s, %s, %s, now());"""
    #         cursor.executemany(query, arr_1)
    #         self.conn.commit()
    #         self.conn.close()

    #     except Exception as e:
    #         error = ("Error", str(e))
    #         return error

    #     return ("완료", "제품위치 정보가 정상적으로 업로드 되었습니다.")

    # def insert_barcode(self, arr_1):
    #     cursor = self.conn.cursor()

    #     try:
    #         query = "TRUNCATE TABLE item_barcode;"
    #         cursor.execute(query)
    #         self.conn.commit()

    #         query = """INSERT INTO item_barcode (item_id, item_name, barcode_in, barcode_out, c_date) VALUES (%s, %s, %s, %s, now());"""
    #         cursor.executemany(query, arr_1)
    #         self.conn.commit()
    #         self.conn.close()            
                        
    #     except Exception as e:
    #         error = ("Error", str(e))
    #         return error

    #     return ("완료", "바코드 정보가 정상적으로 업로드 되었습니다.")

    # def insert_saleslist(self, arr_1):
    #     cursor = self.conn.cursor()

    #     try:
    #         query = "TRUNCATE TABLE sales_list;"
    #         cursor.execute(query)
    #         self.conn.commit()

    #         query = """INSERT INTO sales_list VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,%s, %s, %s, %s, %s, %s, %s, %s, now(), '');"""
    #         cursor.executemany(query, arr_1)
    #         self.conn.commit()
    #         self.conn.close()            
                        
    #     except Exception as e:
    #         error = ("Error", str(e))
    #         return error

    #     return ("완료", "제품 출고정보가 정상적으로 업로드 되었습니다.")

    def insert_overtime(self, arr):
        cursor = self.conn.cursor()

        try:
            query = """INSERT INTO overtime_upload (dept_id, dept_name, emp_id, emp_name, overtime_date, s_time, t_time, overtime, detail, note, c_date) 
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, now());"""
            cursor.executemany(query, arr)
            self.conn.commit()
            self.conn.close()

        except Exception as e:
            error = ("Error", str(e))
            return error

        return ("완료", "잔업 정보가 업로드 되었습니다.")
    
    def update_emp_info(self, arr):
        cursor = self.conn.cursor()

        try:
            query = """
                    UPDATE employee
                    SET dept_id = %s, yn = %s, u_date = NOW()
                    WHERE emp_id = %s;
                    """
            cursor.executemany(query, arr)
            self.conn.commit()
            self.conn.close()

        except Exception as e:
            error = ("Error", str(e))
            return error

        return ("완료", "인사정보가 업데이트 되었습니다.")
    
    def insert_emp_info(self, arr):
        cursor = self.conn.cursor()

        try:
            query = """
                    INSERT INTO employee (dept_id, dept_name, emp_id, emp_name, yn, c_date)
                    VALUES (%s, %s, %s, %s, %s, NOW());
                    """
            cursor.executemany(query, arr)
            self.conn.commit()
            self.conn.close()

        except Exception as e:
            error = ("Error", str(e))
            return error

        return ("완료", "인사정보가 입력 되었습니다.")

# (send_date, order_num, order_date, customer_num, 
#             order_customer, item_id, item_name, serial, ea, warehouse, message_1, destination, send_num, total_amount, second_amount, item_quantity, yesno, inout, item_loc, message_2, box_num, confirm_date, sending_date, message_3, date_1, date_2, ,emp_1, edit_date, completed)