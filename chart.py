import sys
import pandas as pd
import pymysql
import matplotlib.pyplot as plt
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas

class ChartWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Pandas, Matplotlib, PyMySQL and PyQt5 Example")
        self.setGeometry(100, 100, 800, 600)

        # Matplotlib Figure
        self.figure, self.ax = plt.subplots()
        self.canvas = FigureCanvas(self.figure)

        # 버튼
        self.button = QPushButton('Show Chart')
        self.button.clicked.connect(self.plot)

        # 레이아웃
        layout = QVBoxLayout()
        layout.addWidget(self.canvas)
        layout.addWidget(self.button)

        # 위젯 설정
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def get_data_from_db(self):
        # MySQL 데이터베이스 연결 설정
        connection = pymysql.connect(host = "106.248.19.234", port = 3306, database = "webserver", user = "scott", password = "!Jhlee0809" )

        try:
            with connection.cursor() as cursor:
                # SQL 쿼리 실행
                sql = """SELECT dept_name, sum(overtime) from overtime group by dept_name;""" 
                cursor.execute(sql)
                result = cursor.fetchall()
                # 데이터프레임으로 변환
                df = pd.DataFrame(result)
                return df
        finally:
            connection.close()

    def plot(self):
        self.ax.clear()  # 기존 차트 지우기
        df = self.get_data_from_db()
        df.plot(kind='bar', ax=self.ax, color='skyblue')
        self.ax.set_title('항목별 값')
        self.ax.set_xlabel('항목')
        self.ax.set_ylabel('값')
        self.canvas.draw()  # 캔버스 업데이트

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ChartWindow()
    window.show()
    sys.exit(app.exec_())