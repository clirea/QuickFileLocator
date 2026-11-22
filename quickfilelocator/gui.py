"""PyQt6 の画面と、検索を別スレッドで実行する仕組み。"""

from __future__ import annotations

from bisect import bisect_right
from collections.abc import Callable
from pathlib import Path

from PyQt6.QtCore import Qt, QThread, QUrl, pyqtSignal
from PyQt6.QtGui import (
    QCloseEvent,
    QColor,
    QDesktopServices,
    QKeySequence,
    QResizeEvent,
    QShortcut,
    QTextCharFormat,
    QTextCursor,
)
from PyQt6.QtWidgets import (
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
    QSizePolicy,
    QSplitter,
    QStatusBar,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from quickfilelocator.query import QuerySyntaxError, SearchQuery
from quickfilelocator.readers import UnreadableFileError
from quickfilelocator.search import FileSearcher, SearchDirectoryError, SearchSummary

HIT_BACKGROUND = QColor("#fff59d")
CURRENT_HIT_BACKGROUND = QColor("#ffcc80")
PATH_ROLE = Qt.ItemDataRole.UserRole
RELATIVE_PATH_COLOR = "#888888"
PREVIEW_CONTEXT_LINES = 50


def _line_starts(content: str) -> list[int]:
    """本文の各行先頭の文字位置。空文字でも先頭行の 0 を含む。"""
    starts = [0]
    for index, char in enumerate(content):
        if char == "\n":
            starts.append(index + 1)
    return starts


def open_with_default_app(path: str) -> bool:
    """OSで関連付けられたアプリでファイルを開く。"""
    return QDesktopServices.openUrl(QUrl.fromLocalFile(path))


class ElidedLabel(QLabel):
    """幅に合わせて右省略するラベル。"""

    def __init__(self, text: str = "", parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._full_text = text
        self.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Preferred)
        self.setText(text)

    @property
    def full_text(self) -> str:
        return self._full_text

    def resizeEvent(self, event: QResizeEvent | None) -> None:
        elided = self.fontMetrics().elidedText(self._full_text, Qt.TextElideMode.ElideRight, max(self.width(), 0))
        self.setText(elided)
        super().resizeEvent(event)


class ResultItemWidget(QWidget):
    """検索結果1件を、ファイル名・相対パス・抜粋の2段で示す。"""

    def __init__(self, name_text: str, relative_path: str, snippet: str, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 4, 8, 4)
        layout.setSpacing(2)

        row = QHBoxLayout()
        row.setContentsMargins(0, 0, 0, 0)
        self.name_label = QLabel(name_text, self)
        self.name_label.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        self.path_label = QLabel(relative_path, self)
        self.path_label.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        self.path_label.setStyleSheet(f"color: {RELATIVE_PATH_COLOR};")
        row.addWidget(self.name_label)
        row.addStretch(1)
        row.addWidget(self.path_label)

        self.snippet_label = ElidedLabel(snippet, self)
        self.snippet_label.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)

        layout.addLayout(row)
        layout.addWidget(self.snippet_label)


class SearchThread(QThread):
    progress_update = pyqtSignal(int, int)  # 処理済み件数、対象件数
    result_found = pyqtSignal(str, int, str)  # パス、ファイル内ヒット件数、抜粋
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
                on_match=lambda path, hit_count, snippet: self.result_found.emit(str(path), hit_count, snippet),
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
        self._active_query: SearchQuery | None = None
        self._preview_full_content = ""
        self._preview_line_starts: list[int] = [0]
        self._preview_start_line = 0
        self._preview_end_line = 0
        self._hit_spans: list[tuple[int, int]] = []
        self._hit_index = 0
        self.search_thread: SearchThread | None = None

        self.setWindowTitle("ファイル検索")
        self.setGeometry(100, 100, 1100, 700)

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
        self.result_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.result_list.setToolTip(
            "クリックでプレビュー、ダブルクリックでファイルを開きます。選択してコピーもできます。"
        )
        self.result_list.itemActivated.connect(self.open_result)
        self.result_list.currentItemChanged.connect(self.preview_current_result)
        copy_shortcut = QShortcut(QKeySequence.StandardKey.Copy, self.result_list)
        copy_shortcut.activated.connect(self.copy_selected_results)

        results_pane = QWidget(self)
        results_layout = QVBoxLayout(results_pane)
        results_layout.setContentsMargins(0, 0, 0, 0)
        results_layout.addWidget(self.result_list)

        preview_pane = QWidget(self)
        preview_layout = QVBoxLayout(preview_pane)
        preview_layout.setContentsMargins(0, 0, 0, 0)
        preview_header = QHBoxLayout()
        self.preview_label = QLabel("プレビュー", self)
        self.previous_hit_button = QPushButton("前へ", self)
        self.next_hit_button = QPushButton("次へ", self)
        self.previous_hit_button.setEnabled(False)
        self.next_hit_button.setEnabled(False)
        self.previous_hit_button.clicked.connect(self.previous_hit)
        self.next_hit_button.clicked.connect(self.next_hit)
        preview_header.addWidget(self.preview_label)
        preview_header.addStretch(1)
        preview_header.addWidget(self.previous_hit_button)
        preview_header.addWidget(self.next_hit_button)
        self.load_more_above_button = QPushButton("上をもっと読む", self)
        self.load_more_below_button = QPushButton("下をもっと読む", self)
        self.load_more_above_button.setEnabled(False)
        self.load_more_below_button.setEnabled(False)
        self.load_more_above_button.clicked.connect(self.load_more_above)
        self.load_more_below_button.clicked.connect(self.load_more_below)
        self.preview_area = QTextEdit(self)
        self.preview_area.setReadOnly(True)
        self.preview_area.setPlaceholderText("検索結果を選ぶと内容を表示します")
        preview_layout.addLayout(preview_header)
        preview_layout.addWidget(self.load_more_above_button)
        preview_layout.addWidget(self.preview_area)
        preview_layout.addWidget(self.load_more_below_button)
        next_hit_shortcut = QShortcut(QKeySequence("F3"), self)
        next_hit_shortcut.activated.connect(self.next_hit)
        previous_hit_shortcut = QShortcut(QKeySequence("Shift+F3"), self)
        previous_hit_shortcut.activated.connect(self.previous_hit)

        self.results_splitter = QSplitter(Qt.Orientation.Horizontal, self)
        self.results_splitter.addWidget(results_pane)
        self.results_splitter.addWidget(preview_pane)
        self.results_splitter.setStretchFactor(0, 1)
        self.results_splitter.setStretchFactor(1, 1)
        self.results_splitter.setSizes([520, 520])
        layout.addWidget(self.results_splitter, 3)

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
        self._active_query = query
        self.result_list.clear()
        self._clear_preview()
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

    def display_result(self, result: str, hit_count: int = 1, snippet: str = "") -> None:
        path = Path(result)
        item = QListWidgetItem()
        item.setData(PATH_ROLE, result)
        widget = ResultItemWidget(
            f"{path.name}  ({hit_count}件)",
            self._relative_display_path(path),
            snippet,
            self.result_list,
        )
        self.result_list.addItem(item)
        self.result_list.setItemWidget(item, widget)
        item.setSizeHint(widget.sizeHint())
        if self.result_list.currentItem() is None:
            self.result_list.setCurrentRow(0)

    def _relative_display_path(self, path: Path) -> str:
        try:
            directory = Path(self.directory_input.text().strip()).expanduser()
            return str(path.relative_to(directory))
        except ValueError:
            return str(path)

    def preview_current_result(
        self,
        current: QListWidgetItem | None,
        _previous: QListWidgetItem | None = None,
    ) -> None:
        if current is None:
            self._clear_preview()
            return
        path = Path(self._item_path(current))
        try:
            content = self._searcher.read(path)
        except (UnreadableFileError, OSError) as error:
            self._clear_preview()
            self.preview_area.setPlainText(f"プレビューできません: {error}")
            return
        self._preview_full_content = content
        self._preview_line_starts = _line_starts(content)
        query = self._query_for_preview()
        self._hit_spans = list(query.find_matches(content)) if query is not None else []
        self._hit_index = 0
        self._set_initial_preview_window()
        self._render_preview_window()

    def load_more_above(self) -> None:
        if self._preview_start_line <= 0:
            return
        self._preview_start_line = max(0, self._preview_start_line - PREVIEW_CONTEXT_LINES)
        self._render_preview_window()

    def load_more_below(self) -> None:
        last_line = len(self._preview_line_starts) - 1
        if self._preview_end_line >= last_line:
            return
        self._preview_end_line = min(last_line, self._preview_end_line + PREVIEW_CONTEXT_LINES)
        self._render_preview_window()

    def next_hit(self) -> None:
        if len(self._hit_spans) < 2:
            return
        self._hit_index = (self._hit_index + 1) % len(self._hit_spans)
        self._reveal_current_hit()

    def previous_hit(self) -> None:
        if len(self._hit_spans) < 2:
            return
        self._hit_index = (self._hit_index - 1) % len(self._hit_spans)
        self._reveal_current_hit()

    def _item_path(self, item: QListWidgetItem) -> str:
        stored = item.data(PATH_ROLE)
        return stored if stored else item.text()

    def _query_for_preview(self) -> SearchQuery | None:
        if self._active_query is not None:
            return self._active_query
        try:
            return SearchQuery.parse(self.search_input.text(), case_sensitive=self.case_checkbox.isChecked())
        except QuerySyntaxError:
            return None

    def _set_initial_preview_window(self) -> None:
        last_line = len(self._preview_line_starts) - 1
        if self._hit_spans:
            hit_line = self._line_index_for_offset(self._hit_spans[0][0])
            self._preview_start_line = max(0, hit_line - PREVIEW_CONTEXT_LINES)
            self._preview_end_line = min(last_line, hit_line + PREVIEW_CONTEXT_LINES)
            return
        self._preview_start_line = 0
        self._preview_end_line = min(last_line, PREVIEW_CONTEXT_LINES * 2)

    def _visible_offsets(self) -> tuple[int, int]:
        starts = self._preview_line_starts
        start = starts[self._preview_start_line]
        if self._preview_end_line + 1 < len(starts):
            return start, starts[self._preview_end_line + 1]
        return start, len(self._preview_full_content)

    def _line_index_for_offset(self, offset: int) -> int:
        index = bisect_right(self._preview_line_starts, offset) - 1
        return max(0, min(index, len(self._preview_line_starts) - 1))

    def _span_overlaps_visible(self, span: tuple[int, int]) -> bool:
        visible_start, visible_end = self._visible_offsets()
        return span[1] > visible_start and span[0] < visible_end

    def _recenter_window_on_offset(self, offset: int) -> None:
        last_line = len(self._preview_line_starts) - 1
        hit_line = self._line_index_for_offset(offset)
        self._preview_start_line = max(0, hit_line - PREVIEW_CONTEXT_LINES)
        self._preview_end_line = min(last_line, hit_line + PREVIEW_CONTEXT_LINES)

    def _widget_cursor_for_span(self, span: tuple[int, int]) -> QTextCursor:
        visible_start, visible_end = self._visible_offsets()
        start = max(span[0], visible_start) - visible_start
        end = min(span[1], visible_end) - visible_start
        cursor = QTextCursor(self.preview_area.document())
        cursor.setPosition(start)
        cursor.setPosition(end, QTextCursor.MoveMode.KeepAnchor)
        return cursor

    def _render_preview_window(self) -> None:
        visible_start, visible_end = self._visible_offsets()
        self.preview_area.setPlainText(self._preview_full_content[visible_start:visible_end])
        self._apply_hit_decorations()
        self._update_load_more_buttons()

    def _apply_hit_decorations(self) -> None:
        if not self._hit_spans:
            self.preview_area.setExtraSelections([])
            self.preview_label.setText("プレビュー")
            self.next_hit_button.setEnabled(False)
            self.previous_hit_button.setEnabled(False)
            return
        other_format = QTextCharFormat()
        other_format.setBackground(HIT_BACKGROUND)
        current_format = QTextCharFormat()
        current_format.setBackground(CURRENT_HIT_BACKGROUND)
        current_span = self._hit_spans[self._hit_index]
        extras: list[QTextEdit.ExtraSelection] = []
        for span in self._hit_spans:
            if not self._span_overlaps_visible(span):
                continue
            selection = QTextEdit.ExtraSelection()
            selection.cursor = self._widget_cursor_for_span(span)
            selection.format = current_format if span == current_span else other_format
            extras.append(selection)
        self.preview_area.setExtraSelections(extras)
        if self._span_overlaps_visible(current_span):
            current = self._widget_cursor_for_span(current_span)
            self.preview_area.setTextCursor(current)
            self.preview_area.ensureCursorVisible()
        total = len(self._hit_spans)
        self.preview_label.setText(f"プレビュー  {self._hit_index + 1} / {total}")
        movable = total > 1
        self.next_hit_button.setEnabled(movable)
        self.previous_hit_button.setEnabled(movable)

    def _update_load_more_buttons(self) -> None:
        last_line = len(self._preview_line_starts) - 1
        self.load_more_above_button.setEnabled(self._preview_start_line > 0)
        self.load_more_below_button.setEnabled(self._preview_end_line < last_line)

    def _reveal_current_hit(self) -> None:
        current_span = self._hit_spans[self._hit_index]
        if not self._span_overlaps_visible(current_span):
            self._recenter_window_on_offset(current_span[0])
        self._render_preview_window()

    def _clear_preview(self) -> None:
        self._preview_full_content = ""
        self._preview_line_starts = [0]
        self._preview_start_line = 0
        self._preview_end_line = 0
        self._hit_spans = []
        self._hit_index = 0
        self.preview_area.setExtraSelections([])
        self.preview_area.clear()
        self.preview_label.setText("プレビュー")
        self.next_hit_button.setEnabled(False)
        self.previous_hit_button.setEnabled(False)
        self.load_more_above_button.setEnabled(False)
        self.load_more_below_button.setEnabled(False)

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
        path = self._item_path(item)
        if not self._open_path(path):
            self.status_bar.showMessage(f"ファイルを開けませんでした: {path}")

    def copy_selected_results(self) -> None:
        paths = [self._item_path(item) for item in self.result_list.selectedItems()]
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
