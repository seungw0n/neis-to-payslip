import subprocess
import sys
from src import excel, fill, form, neis, payslip
from PyQt5 import uic, QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import *

form_class = uic.loadUiType("main.ui")[0]


class WindowClass(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.initUI()
        self.neis_obj = None
        self.form_obj = None

    def initUI(self):
        self.setWindowTitle("임금명세서 - Beta version")
        self.selectNeis.clicked.connect(self.neis_button)
        self.selectForm.clicked.connect(self.form_button)
        self.createPayslip.clicked.connect(self.payslip_button)

    def neis_button(self):
        filepath = QFileDialog.getOpenFileName(self, "Select a Excel File", filter="*.xlsx")
        print("Filepath : " + str(filepath))

        if filepath != ('', ''):
            filename = filepath[0].split("/")  # comes only filename
            self.labelNeisPath.setText(filename[-1])

            neis_wb = excel.open_excel(filepath[0])
            try:
                self.neis_obj = neis.NeisPayslip(neis_wb)
                self.neis_obj.get_pay()
                self.neis_obj.get_tax()
                self.neis_obj.get_deduction()

                self.labelSchoolName.setText(self.neis_obj.school_name)
                self.labelEmployeeName.setText(self.neis_obj.employee_name)
                self.labelPosition.setText(self.neis_obj.position)
            except ValueError:
                self.warning("Neis 급여명세서가 아닙니다.")
                self.reset()
        else:
            return

    def form_button(self):
        if self.labelSchoolName.text() == "":
            self.warning("Neis 임금명세서를 찾을 수 없습니다.")
            return

        filepath = QFileDialog.getOpenFileName(self, "Select a Excel File", filter="*.xlsx")
        print("Filepath : " + str(filepath))

        if filepath != ('', ''):
            filename = filepath[0].split("/")
            self.labelFormPath.setText(filename[-1])

            form_wb = excel.open_excel(filepath[0])
            try:
                self.form_obj = form.Form(form_wb)
                self.form_obj.get_info()
                print(self.form_obj.info)

                if self.form_obj.info['성명'] != self.neis_obj.employee_name:
                    self.warning("성명이 일치하지 않습니다.")
                    self.reset(all=False)
            except ValueError as e:
                self.warning("제공된 엑셀양식이 아닙니다.")
                self.reset(all=False)
        else:
            return

    def warning(self, message):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)
        msg.setText(message)
        # msg.setInformativeText("Try again")
        msg.setWindowTitle("Error")
        msg.exec()

    def reset(self, all=True):
        self.form_obj = None
        self.labelFormPath.setText("")
        if all:
            self.neis_obj = None
            self.labelNeisPath.setText("")
            self.labelSchoolName.setText("")
            self.labelEmployeeName.setText("")
            self.labelPosition.setText("")

    def payslip_button(self):
        school_obj = payslip.SchoolPayslip(self.neis_obj, self.form_obj)
        school_obj.match_pay()
        school_obj.match_tax()
        school_obj.match_deduction()
        school_obj.match_total()
        print(school_obj.total)
        #         top_field(sample_sheet, neis_obj, form_obj)
        #         monthly_field(sample_sheet, school_obj)
        #         non_periodically_field(sample_sheet, school_obj)
        #         rest_field(sample_sheet, school_obj)
        #         deduction_field(sample_sheet, school_obj)
        #         total_field(sample_sheet, school_obj)

        out_wb = excel.open_excel("payslip.xlsx")
        out_sheet = out_wb['Sheet1']
        fill.fill(out_sheet, neis_obj=self.neis_obj, form_obj=self.form_obj, school_obj=school_obj)
        try:
            excel.save_excel(out_wb, "./임금명세서_" + self.neis_obj.employee_name + ".xlsx")
        except Exception as e:
            print(e)

        self.reset()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = WindowClass()
    myWindow.show()
    app.exec_()
