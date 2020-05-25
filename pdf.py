import io
from PyPDF2 import PdfFileWriter, PdfFileReader
from pdfminer.converter import TextConverter
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfpage import PDFPage
import pandas as pd
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication, QWidget, QInputDialog, QLineEdit, QFileDialog
from PyQt5.QtGui import QIcon
import sys
import os
import time
import xlsxwriter

initial,used, used_more, x = [],[],[],[]


class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(400, 250)
        self.msg1 = QtWidgets.QLabel(Form)
        self.msg2 = QtWidgets.QLabel(Form)
        self.msg1.setGeometry(QtCore.QRect(15, 10, 300, 60))
        self.msg1.setText('')
        self.msg2.setGeometry(QtCore.QRect(15,60,350,60))
        self.msg2.setText('')
        self.pushButton = QtWidgets.QPushButton(Form)
        self.pushButton.setGeometry(QtCore.QRect(260, 160, 125, 32))#PDF path
        self.pushButton.setStyleSheet("background-color:red;\n"
                                      "color: white;\n"
                                      "border-style: outset;\n"
                                      "border-width:2px;\n"
                                      "border-radius:10px;\n"
                                      "border-color:black;\n"
                                      "font:bold 14px;\n"
                                      "padding :6px;\n"
                                      "min-width:10px;\n"
                                      "\n"
                                      "\n"
                                      "")

        self.pushButton.setObjectName("pushButton")
        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "File shuffler"))
        self.pushButton.setText(_translate("Form", "Browse PDF"))
        self.pushButton.clicked.connect(self.open_dialog_box)


    def open_dialog_box(self, Form):
        filename = QFileDialog.getOpenFileName()
        pdf_path = filename[0]
        print(pdf_path)
        if pdf_path == "":
            self.msg2.setText("Select a PDF file to execute: ")
        else:
            self.msg2.setText("Waiting")
            self.extract_text(pdf_path)
            self.post_execute(pdf_path)
            skipped = ""
            for item in initial:
                skipped += str(item) +","
            self.msg1.setText('<h1>Execution completed</h1>')
            self.msg2.setText('Pages skipped:'+ skipped)

    def extract_text_by_page(self, pdf_path):
        if os.path.exists(pdf_path):
            try:
                with open(pdf_path, 'rb') as fh:
                    page_count = 1
                    initial.clear()
                    used.clear()
                    used_more.clear()
                    for page in PDFPage.get_pages(fh, caching=True, check_extractable=True):
                        if page_count % 2 == 0:
                            initial.append(page_count)
                        resource_manager = PDFResourceManager()
                        fake_file_handle = io.StringIO()
                        converter = TextConverter(resource_manager, fake_file_handle)
                        page_interpreter = PDFPageInterpreter(resource_manager, converter)
                        page_interpreter.process_page(page)
                        text = fake_file_handle.getvalue()
                        yield text
                        # close open handles
                        converter.close()
                        fake_file_handle.close()
                        page_count += 1
                fh.close()
            except:
                self.msg2.setText("PDF File not found")


    def extract_text(self, pdf_path):
        with open('out.txt', 'w+') as f:
            for page in self.extract_text_by_page(pdf_path):
                print(page.encode("utf-8"), file=f)
        f.close()

    def find_String(self, str_line,digit,name, slot):
        amt_parse = str_line.find("TotalAmount")
        if amt_parse != -1:
            name = str(str_line[amt_parse+12:slot-2])
        qty_parse = str_line.find("\\xe")

        if qty_parse != -1:
            quantity = str(str_line[qty_parse + 18:qty_parse + 20])
            for value in quantity:
                if value.isdigit():
                    digit += value
        return digit, name

    def post_execute(self, pdf_path):
        file_name = os.path.splitext(os.path.split(pdf_path)[1])[0] + "_" + time.strftime("%m-%d-%H%M%S") #parsing the input file from path
        out_file = file_name + ".pdf"
        xl_file = file_name + ".xlsx"
        path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop\Out')
        out_pdf_path = os.path.join(path, out_file)
        out_xl_path = os.path.join(path, xl_file)
        if not os.path.exists(path):
            os.makedirs(path)
        output_pdf = PdfFileWriter()
        if os.path.exists("ASIN.xlsx"):
            try:
                data = pd.read_excel("ASIN.xlsx",header=None)
            except:
                self.msg2.setText('ASIN file not found')
        ASIN = data.iloc[:,0].values.tolist()
        length = max(initial)
        workbook = xlsxwriter.Workbook(out_xl_path)
        worksheet = workbook.add_worksheet()

        row, col = 0, 0
        product_name, qty = '', ''
        for val in ASIN:
            total_qty = 0
            val_app = val.strip()
            with open('out.txt', 'r') as d:
                for page in range(1,length+1):
                    digit, name = '', ''
                    str_line = d.readline()
                    slot = str_line.find(str(val_app))
                    if slot != -1:
                        qty, product_name = self.find_String(str_line, digit, name, slot)

                        if page in initial and page not in used and page not in used_more:
                            if int(qty) > 1:
                                used_more.append(page)
                            else:
                                used.append(page)
                        try:
                            initial.remove(page)
                        except:
                            pass
                        total_qty += int(qty)

                if total_qty > 0:
                    worksheet.write(row, col, product_name)
                    worksheet.write(row, col+1, val_app)
                    worksheet.write(row, col + 2, total_qty)
                    row+=1
        d.close()
        workbook.close()
        temp = initial
        print("Reshuffled",used)
        print("More Quatity",used_more)
        print("Pending",temp)

        with open(pdf_path,'rb') as readfile:
            input_pdf = PdfFileReader(readfile, strict=None)
            for page in used_more:
                output_pdf.addPage(input_pdf.getPage(page-2))
                output_pdf.addPage(input_pdf.getPage(page-1))
            for page in used:
                output_pdf.addPage(input_pdf.getPage(page-2))
                output_pdf.addPage(input_pdf.getPage(page-1))
            with open(out_pdf_path, "wb") as writefile:
                output_pdf.write(writefile)
        readfile.close()


def main():
    app = QtWidgets.QApplication(sys.argv)
    Form = QtWidgets.QWidget()
    ui = Ui_Form()
    ui.setupUi(Form)
    Form.show()
    sys.exit(app.exec_())

main()