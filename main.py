import os
import sys
import chardet
import re
import docx
from PyQt5.QtWidgets import QApplication, QMainWindow, QTextEdit, QLineEdit, QPushButton, QLabel, QVBoxLayout, QHBoxLayout, QWidget, QProgressBar, QStatusBar
from PyQt5.QtCore import QThread, pyqtSignal

class SearchThread(QThread):
    progress_update = pyqtSignal(int)
    result_found = pyqtSignal(str)
    search_completed = pyqtSignal(int, int, list)  # 検索したファイル数、除外したファイル数、除外したファイル名を渡すシグナル

    def __init__(self, search_string):
        super().__init__()
        self.search_string = search_string
        self.files_searched = 0  # 検索したファイル数を保持する属性
        self.files_excluded = 0  # 除外したファイル数を保持する属性
        self.excluded_files = []  # 除外したファイル名を保持する属性

    def run(self):
        self.search_files(self.search_string, os.path.dirname(os.path.abspath(__file__)))
        self.search_completed.emit(self.files_searched, self.files_excluded, self.excluded_files)  # 検索したファイル数、除外したファイル数、除外したファイル名を渡す

    def search_files(self, search_string, directory):
        extensions = [".docx", ".txt"]

        # 検索文字列を解析
        and_parts = re.split(r'\s*&&\s*', search_string)
        or_parts = []
        for part in and_parts:
            or_parts.append(re.split(r'\s*\|\|\s*', part))

        total_files = sum([len(files) for _, _, files in os.walk(directory)])
        processed_files = 0

        for root, dirs, files in os.walk(directory):
            for file in files:
                if file.endswith(tuple(extensions)):
                    file_path = os.path.join(root, file)
                    try:
                        if file.endswith(".docx"):
                            if file.startswith("~$"):  # 開いているファイルを除外
                                self.files_excluded += 1
                                self.excluded_files.append(file_path)
                                continue
                            doc = docx.Document(file_path)
                            content = "\n".join([para.text for para in doc.paragraphs])
                        else:
                            with open(file_path, "rb") as f:
                                content_bytes = f.read()
                            encoding = chardet.detect(content_bytes)["encoding"]
                            try:
                                content = content_bytes.decode(encoding)
                            except UnicodeDecodeError:
                                # デコードエラーが発生した場合は除外
                                self.files_excluded += 1
                                self.excluded_files.append(file_path)
                                continue

                        # アンド検索とオア検索の条件を確認
                        match = True
                        for and_part in or_parts:
                            or_match = False
                            for or_part in and_part:
                                if re.search(r'\s+', or_part):
                                    # 引用符で囲まれたフレーズ検索
                                    if re.search(r'"\s*' + re.escape(or_part) + r'\s*"', content):
                                        or_match = True
                                        break
                                else:
                                    # 単語検索
                                    if or_part in content:
                                        or_match = True
                                        break
                            if not or_match:
                                match = False
                                break

                        if match:
                            self.result_found.emit(file_path)
                    except (docx.opc.exceptions.PackageNotFoundError, IOError):
                        self.files_excluded += 1
                        self.excluded_files.append(file_path)

                processed_files += 1
                self.files_searched += 1  # 検索したファイル数をカウント
                progress = int((processed_files / total_files) * 100)
                self.progress_update.emit(progress)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ファイル検索")
        self.setGeometry(100, 100, 800, 600)

        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)

        layout = QVBoxLayout(central_widget)

        input_layout = QHBoxLayout()
        self.search_input = QLineEdit(self)
        self.search_input.setPlaceholderText("検索文字列を入力してください(or検索(空白か||)、and検索(""か&&)が可能)")
        input_layout.addWidget(self.search_input)

        layout.addLayout(input_layout)

        self.search_button = QPushButton("検索", self)
        self.search_button.clicked.connect(self.start_search)
        layout.addWidget(self.search_button)

        self.result_area = QTextEdit(self)
        self.result_area.setReadOnly(True)
        layout.addWidget(self.result_area)

        self.progress_bar = QProgressBar(self)
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        self.status_bar = QStatusBar(self)
        self.setStatusBar(self.status_bar)

        self.search_thread = None

    def start_search(self):
        search_string = self.search_input.text()
        if search_string:
            self.search_button.setEnabled(False)
            self.result_area.clear()
            self.progress_bar.setVisible(True)
            self.progress_bar.setValue(0)

            self.search_thread = SearchThread(search_string)
            self.search_thread.progress_update.connect(self.update_progress)
            self.search_thread.result_found.connect(self.display_result)
            self.search_thread.search_completed.connect(self.search_completed)
            self.search_thread.start()

    def update_progress(self, value):
        self.progress_bar.setValue(value)
        self.status_bar.showMessage(f"検索進捗: {value}%")

    def display_result(self, result):
        self.result_area.append(result)

    def search_completed(self, files_searched, files_excluded, excluded_files):
        self.search_button.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.status_bar.showMessage(f"検索完了。 {files_searched} 件のファイルを検索し、 {files_excluded} 件のファイルを除外しました。")  # 検索したファイル数と除外したファイル数を表示

        # 除外したファイル名を結果エリアに表示
        self.result_area.append("\n除外されたファイル:")
        for file in excluded_files:
            self.result_area.append(file)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec_())
