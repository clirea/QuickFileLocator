"""PyQt5 の画面と、検索を別スレッドで実行する仕組み。"""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path

from PyQt5.QtCore import QThread, QUrl, pyqtSignal
from PyQt5.QtGui import QCloseEvent, QDesktopServices, QKeySequence
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QShortcut,
    QStatusBar,
    QVBoxLayout,
    QWidget,
)

from quickfilelocator.query import QuerySyntaxError, SearchQuery
from quickfilelocator.search import FileSearcher, SearchDirectoryError, SearchSummary


def open_with_default_app(path: str) -> bool:
    """OSで関連付けられたアプリでファイルを開く。"""
    return QDesktopServices.openUrl(QUrl.fromLocalFile(path))


class SearchThread(QThread):
    progress_update = pyqtSignal(int, int)  # 処理済み件数、対象件数
    result_found = pyqtSignal(str)
    search_completed = pyqtSignal(object)  # SearchSummary
    search_failed = pyqtSignal(str)

    def __init__(self, searcher: FileSearcher, query: SearchQuery, directory: Path) -> None:
        super().__init__()
        self._searcher = searcher
        self._query = query
        self._directory = directory
        self._last_percent = -1

    def run(self) -> None:
        # ここから例外が漏れるとQtがプロセスごと異常終了させるため、必ず画面へ通知する
        try:
            summary = self._searcher.search(
                self._query,
                self._directory,
                on_match=lambda path: self.result_found.emit(str(path)),
                on_progress=self._report_progress,
                is_cancelled=self.isInterruptionRequested,
            )
        except SearchDirectoryError as error:
            self.search_failed.emit(str(error))
        except Exception as error:
            self.search_failed.emit(f"検索中にエラーが発生しました: {type(error).__name__}: {error}")
        else:
            self.search_completed.emit(summary)

    def _report_progress(self, processed_count: int, total_count: int) -> None:
        """ファイル数が多くても画面更新が詰まらないよう、1%刻みでだけ通知する。"""
        percent = processed_count * 100 // total_count if total_count else 100
        if percent != self._last_percent:
            self._last_percent = percent
            self.progress_update.emit(processed_count, total_count)


class MainWindow(QMainWindow):
    def __init__(
        self,
        searcher: FileSearcher,
        initial_directory: Path,
        thread_factory: Callable[[FileSearcher, SearchQuery, Path], SearchThread] = SearchThread,
        open_path: Callable[[str], bool] = open_with_default_app,
    ) -> None:
        super().__init__()
        self._searcher = searcher
        self._thread_factory = thread_factory
        self._open_path = open_path
        self._close_requested = False
        self.search_thread: SearchThread | None = None

        self.setWindowTitle("ファイル検索")
        self.setGeometry(100, 100, 800, 600)

        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)

        layout = QVBoxLayout(central_widget)

        directory_layout = QHBoxLayout()
        directory_layout.addWidget(QLabel("検索先:", self))
        self.directory_input = QLineEdit(str(initial_directory), self)
        self.directory_input.returnPressed.connect(self.start_search)
        directory_layout.addWidget(self.directory_input)
        self.browse_button = QPushButton("参照...", self)
        self.browse_button.clicked.connect(self.choose_directory)
        directory_layout.addWidget(self.browse_button)
        layout.addLayout(directory_layout)

        input_layout = QHBoxLayout()
        self.search_input = QLineEdit(self)
        self.search_input.setPlaceholderText('検索文字列 (OR: 空白か || / AND: && / フレーズ: "..." で囲む)')
        self.search_input.returnPressed.connect(self.start_search)
        input_layout.addWidget(self.search_input)
        self.case_checkbox = QCheckBox("大文字と小文字を区別", self)
        self.case_checkbox.setChecked(True)
        input_layout.addWidget(self.case_checkbox)
        layout.addLayout(input_layout)

        button_layout = QHBoxLayout()
        self.search_button = QPushButton("検索", self)
        self.search_button.clicked.connect(self.start_search)
        button_layout.addWidget(self.search_button)
        self.cancel_button = QPushButton("キャンセル", self)
        self.cancel_button.setEnabled(False)
        self.cancel_button.clicked.connect(self.cancel_search)
        button_layout.addWidget(self.cancel_button)
        layout.addLayout(button_layout)

        self.result_list = QListWidget(self)
        self.result_list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.result_list.setToolTip("ダブルクリックでファイルを開きます。選択してコピーもできます。")
        self.result_list.itemActivated.connect(self.open_result)
        layout.addWidget(self.result_list, 3)
        copy_shortcut = QShortcut(QKeySequence.Copy, self.result_list)
        copy_shortcut.activated.connect(self.copy_selected_results)

        self.excluded_label = QLabel("除外されたファイル:", self)
        self.excluded_label.setVisible(False)
        layout.addWidget(self.excluded_label)
        self.excluded_area = QPlainTextEdit(self)
        self.excluded_area.setReadOnly(True)
        self.excluded_area.setVisible(False)
        layout.addWidget(self.excluded_area, 1)

        self.progress_bar = QProgressBar(self)
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        self.status_bar = QStatusBar(self)
        self.setStatusBar(self.status_bar)

    def choose_directory(self) -> None:
        selected = QFileDialog.getExistingDirectory(self, "検索先フォルダーを選択", self.directory_input.text())
        if selected:
            self.directory_input.setText(selected)

    def start_search(self) -> None:
        if self.search_thread is not None:
            return
        try:
            query = SearchQuery.parse(self.search_input.text(), case_sensitive=self.case_checkbox.isChecked())
        except QuerySyntaxError as error:
            self.status_bar.showMessage(str(error))
            return
        directory_text = self.directory_input.text().strip()
        directory = Path(directory_text).expanduser()
        if not directory_text or not directory.is_dir():
            self.status_bar.showMessage(f"検索先フォルダーが見つかりません: {directory_text}")
            return

        self._set_searching(True)
        self.result_list.clear()
        self.excluded_area.clear()
        self.excluded_label.setVisible(False)
        self.excluded_area.setVisible(False)
        # 対象ファイルの件数が分かるまでは、進捗バーを「処理中」表示にする
        self.progress_bar.setRange(0, 0)
        self.progress_bar.setVisible(True)
        self.status_bar.showMessage("対象ファイルを数えています...")

        self.search_thread = self._thread_factory(self._searcher, query, directory)
        self.search_thread.progress_update.connect(self.update_progress)
        self.search_thread.result_found.connect(self.display_result)
        self.search_thread.search_completed.connect(self.search_completed)
        self.search_thread.search_failed.connect(self.search_failed)
        self.search_thread.finished.connect(self.search_finished)
        self.search_thread.start()

    def cancel_search(self) -> None:
        if self.search_thread is None:
            return
        self.search_thread.requestInterruption()
        self.cancel_button.setEnabled(False)
        self.status_bar.showMessage("キャンセルしています...")

    def update_progress(self, processed_count: int, total_count: int) -> None:
        self.progress_bar.setRange(0, max(total_count, 1))
        self.progress_bar.setValue(processed_count)
        self.status_bar.showMessage(f"検索中: {processed_count} / {total_count} 件")

    def display_result(self, result: str) -> None:
        self.result_list.addItem(result)

    def search_completed(self, summary: SearchSummary) -> None:
        prefix = "検索をキャンセルしました。" if summary.cancelled else "検索完了。"
        self.status_bar.showMessage(
            f"{prefix} 対象 {summary.target_count} 件中 {summary.searched_count} 件を検索し、"
            f"{summary.matched_count} 件が一致、{len(summary.excluded)} 件を除外しました。"
        )
        if summary.excluded:
            self.excluded_area.setPlainText("\n".join(f"{item.path}  [{item.reason}]" for item in summary.excluded))
            self.excluded_label.setVisible(True)
            self.excluded_area.setVisible(True)

    def search_failed(self, message: str) -> None:
        self.status_bar.showMessage(message)

    def search_finished(self) -> None:
        """成功・失敗・キャンセルのどれで終わっても、スレッドを片付けて操作可能な状態に戻す。"""
        if self.search_thread is not None:
            self.search_thread.deleteLater()
            self.search_thread = None
        self._set_searching(False)
        self.progress_bar.setVisible(False)
        if self._close_requested:
            self.close()

    def open_result(self, item: QListWidgetItem) -> None:
        if not self._open_path(item.text()):
            self.status_bar.showMessage(f"ファイルを開けませんでした: {item.text()}")

    def copy_selected_results(self) -> None:
        paths = [item.text() for item in self.result_list.selectedItems()]
        if paths:
            QApplication.clipboard().setText("\n".join(paths))

    def closeEvent(self, event: QCloseEvent) -> None:
        # 実行中のスレッドを残したまま破棄すると異常終了するため、停止を待ってから閉じる
        if self.search_thread is None:
            event.accept()
            return
        self._close_requested = True
        self.cancel_search()
        event.ignore()

    def _set_searching(self, searching: bool) -> None:
        self.search_button.setEnabled(not searching)
        self.browse_button.setEnabled(not searching)
        self.directory_input.setEnabled(not searching)
        self.cancel_button.setEnabled(searching)
