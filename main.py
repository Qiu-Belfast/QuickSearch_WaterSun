import sys
import os
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QPushButton, QLabel, QLineEdit, QVBoxLayout, QHBoxLayout, QComboBox

from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from docx import Document

class Window(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("PyQt Example")
        self.setGeometry(100, 100, 450, 150)

        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        self.EntryInputKeyword = QLineEdit(self)
        self.EntryInputPath = QLineEdit(self)
        self.EntryInputPath.setText(os.getcwd())  # 默认路径为当前工作目录

        layout.addLayout(self.create_label_entry("关键字", self.EntryInputKeyword))
        layout.addLayout(self.create_label_entry("路径", self.EntryInputPath))

        self.select = QComboBox()
        self.select.addItems(["PDF", "Word", "MD"])
        layout.addLayout(self.create_label_entry("导出格式", self.select))

        self.buttonSearch = QPushButton("寻找", self)
        self.buttonExport = QPushButton("导出", self)
        self.buttonQuit = QPushButton("退出", self)

        self.buttonSearch.clicked.connect(self.search_the_file)
        self.buttonExport.clicked.connect(self.export_the_file)
        self.buttonQuit.clicked.connect(self.close)

        button_layout = QHBoxLayout()
        button_layout.addWidget(self.buttonSearch)
        button_layout.addStretch()
        button_layout.addWidget(self.buttonExport)
        button_layout.addWidget(self.buttonQuit)

        layout.addLayout(button_layout)

        self.LabelOutputinstruct = QLabel("该路径下的文件为：")
        self.LabelOutputmessage = QLabel()
        layout.addWidget(self.LabelOutputinstruct)
        layout.addWidget(self.LabelOutputmessage)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def create_label_entry(self, label_text, entry_widget):
        layout = QHBoxLayout()
        label = QLabel(label_text)
        layout.addWidget(label)
        layout.addWidget(entry_widget)
        return layout

    def search_file(self):
        path = self.EntryInputPath.text()
        keyword = self.EntryInputKeyword.text()
        match_files = []

        for root, dirs, files in os.walk(path):
            for file_search in files:
                if keyword in file_search:
                    match_files.append(os.path.join(root, file_search))

        return match_files

    def search_the_file(self):
        self.LabelOutputinstruct.setText("该路径下的文件为：")
        match_files = self.search_file()

        if match_files:
            self.LabelOutputmessage.setText("\n".join(match_files))
        else:
            self.LabelOutputmessage.setText("未找到匹配文件")

    def export_the_file(self):
        flag = self.select.currentText()
        match_files = self.search_file()

        if flag == 'PDF':
            c = canvas.Canvas("output.pdf", pagesize=letter)
            y = 700
            for file_search in match_files:
                c.drawString(100, y, file_search)
                y -= 20
            c.save()
        elif flag == 'Word':
            doc = Document()
            for file_search in match_files:
                doc.add_paragraph(file_search)
            doc.save("output.docx")
        elif flag == 'MD':
            with open("output.md", "w") as file:
                for file_search in match_files:
                    file.write("- " + file_search + "\n")
        else:
            print("error")


app = QApplication(sys.argv)
window = Window()
window.show()
sys.exit(app.exec_())
