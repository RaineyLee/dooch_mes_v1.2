
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QPushButton
import sys

class MainWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        # 예시 위젯 추가
        button1 = QPushButton('Button 1', self)
        button2 = QPushButton('Button 2', self)

        layout.addWidget(button1)
        layout.addWidget(button2)

        self.setLayout(layout)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Main Window with Sub Widget')
        self.setGeometry(100, 100, 600, 400)

        # 메인 윈도우에 위젯 추가
        main_widget = MainWidget()
        self.setCentralWidget(main_widget)

        self.show()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    main_window = MainWindow()
    