import os
import sys
import chardet
import re
import docx
from PyQt6.QtWidgets import (
    QApplication,
    QMainWindow,
    QTextEdit,
    QLineEdit,
    QPushButton,
    QLabel,
    QVBoxLayout,
    QHBoxLayout,
    QWidget,
    QProgressBar,
    QStatusBar,
    QFileDialog,
)
from PyQt6.QtCore import QThread, pyqtSignal

MAX_SEARCH_FILE_BYTES = 32 * 1024 * 1024


class SearchDirectoryError(ValueError):
    """検索対象フォルダの指定が不正な場合の例外。"""


def resolve_search_directory(path):
    stripped = path.strip().strip('"')
    if not stripped:
        raise SearchDirectoryError("検索対象フォルダを指定してください。")
    if not os.path.isdir(stripped):
        raise SearchDirectoryError("指定されたフォルダが存在しません。")
    return os.path.abspath(stripped)


class SearchThread(QThread):
    progress_update = pyqtSignal(int)
    result_found = pyqtSignal(str)
    search_completed = pyqtSignal(int, int, list)  # 検索したファイル数、除外したファイル数、除外したファイル名を渡すシグナル

    def __init__(self, search_string, search_directory):
        super().__init__()
        self.search_string = search_string
        self.search_directory = search_directory
        self.files_searched = 0  # 検索したファイル数を保持する属性
        self.files_excluded = 0  # 除外したファイル数を保持する属性
        self.excluded_files = []  # 除外したファイル名を保持する属性

    def run(self):
        self.search_files(self.search_string, self.search_directory)
        self.search_completed.emit(self.files_searched, self.files_excluded, self.excluded_files)  # 検索したファイル数、除外したファイル数、除外したファイル名を渡す

    def _exclude_file(self, file_path):
        self.files_excluded += 1
        self.excluded_files.append(file_path)

    def search_files(self, search_string, directory):
        extensions = [".docx", ".txt"]

        # 検索文字列を解析
        and_parts = re.split(r'\s*&&\s*', search_string)
        or_parts = []
        for part in and_parts:
            or_parts.append(re.split(r'\s*\|\|\s*', part))

        total_files = sum([len(files) for _, _, files in os.walk(directory)])
        processed_files = 0
        if total_files == 0:
            self.progress_update.emit(100)

        for root, dirs, files in os.walk(directory):
            for file in files:
                if file.endswith(tuple(extensions)):
                    file_path = os.path.join(root, file)
                    try:
                        try:
                            if os.path.getsize(file_path) > MAX_SEARCH_FILE_BYTES:
                                self._exclude_file(file_path)
                                continue
                        except OSError:
                            pass

                        if file.endswith(".docx"):
                            if file.startswith("~$"):  # 開いているファイルを除外
                                self._exclude_file(file_path)
                                continue
                            doc = docx.Document(file_path)
                            content = "\n".join([para.text for para in doc.paragraphs])
                        else:
                            with open(file_path, "rb") as f:
                                content_bytes = f.read()
                            encoding = chardet.detect(content_bytes)["encoding"]
                            if not encoding:
                                self._exclude_file(file_path)
                                continue
                            try:
                                content = content_bytes.decode(encoding)
                            except (UnicodeDecodeError, LookupError):
                                self._exclude_file(file_path)
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
                    except Exception:
                        self._exclude_file(file_path)

                processed_files += 1
                self.files_searched += 1  # 検索したファイル数をカウント
                if total_files > 0:
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

        folder_layout = QHBoxLayout()
        folder_label = QLabel("検索対象フォルダ:", self)
        self.folder_input = QLineEdit(self)
        self.folder_input.setPlaceholderText("検索対象フォルダのパスを入力してください")
        self.browse_button = QPushButton("参照", self)
        self.browse_button.clicked.connect(self.browse_directory)
        folder_layout.addWidget(folder_label)
        folder_layout.addWidget(self.folder_input)
        folder_layout.addWidget(self.browse_button)
        layout.addLayout(folder_layout)

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

    def browse_directory(self):
        current = self.folder_input.text().strip().strip('"')
        start_dir = current if os.path.isdir(current) else os.path.expanduser("~")
        directory = QFileDialog.getExistingDirectory(self, "検索対象フォルダを選択", start_dir)
        if directory:
            self.folder_input.setText(os.path.normpath(directory))

    def start_search(self):
        search_string = self.search_input.text()
        if not search_string:
            self.status_bar.showMessage("検索文字列を入力してください。")
            return

        try:
            search_directory = resolve_search_directory(self.folder_input.text())
        except SearchDirectoryError as error:
            self.status_bar.showMessage(str(error))
            return

        self.search_button.setEnabled(False)
        self.browse_button.setEnabled(False)
        self.folder_input.setEnabled(False)
        self.result_area.clear()
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)

        self.search_thread = SearchThread(search_string, search_directory)
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
        self.browse_button.setEnabled(True)
        self.folder_input.setEnabled(True)
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
    sys.exit(app.exec())
