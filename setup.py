import sys
from cx_Freeze import setup, Executable
from PyQt5.QtWidgets import QApplication, QMessageBox
import os

try:
    # 개발 참조 사이트
    # https://cx-freeze.readthedocs.io/en/latest/setup_script.html#cmdoption-arg-include_files

    packages = [
        "PyQt5.QtWidgets"
    ]

    # Output directory
    build_exe_options = {
        "packages": packages,
        "excludes": [],
        "include_files": [("./ui", "ui"), 
                          ("./excel", "excel"),
                          ("./utils", "utils"),
                          ("./popup", "popup"), 
                          ("./db", "db"), 
                          (r"C:/Python310/python310.dll", "python310.dll")
                          ],  # DLL 파일 추가 # 특정 폴더 추가가 필요시 경로명(상대경로 가능)과 사용할 폴더명 명시
        "build_exe": "D:/2. myproject/Build"  # Specify the output directory
    }

    setup(
        name="DOOCHPUMP MES_V1.0",
        version="1.0",
        description="DOOCHPUMP MES",
        executables=[
            Executable(
                "main.py",
                copyright="Copyright (C) 2023 cx_Freeze",
                # base="C:\Python Workplace\Make_ERP",
                icon="warehouse.ico",
            ),
        ],
        options={
            "build_exe": build_exe_options,
            # "bdist_msi": bdist_msi_options,
        },
    )
except Exception as e:
    app = QApplication(sys.argv)
    msg = QMessageBox()
    msg.setWindowTitle("Error")               # 제목설정
    msg.setText(str(e))                          # 내용설정
    msg.exec_()    