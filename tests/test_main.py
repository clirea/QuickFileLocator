import os
from unittest.mock import MagicMock, patch

import pytest
from PyQt6.QtWidgets import QApplication

from main import (
    MAX_SEARCH_FILE_BYTES,
    MainWindow,
    SearchDirectoryError,
    SearchThread,
    resolve_search_directory,
)


@pytest.fixture(scope="session")
def qapp():
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


@pytest.fixture
def window(qapp):
    main_window = MainWindow()
    yield main_window
    main_window.close()


@pytest.fixture
def search_thread(qapp):
    return SearchThread("needle", r"C:\docs")


@patch("main.os.path.isdir", return_value=True)
@patch("main.os.path.abspath", return_value=r"C:\docs")
def test正常系_resolve_search_directoryメソッドが存在するフォルダの絶対パスを返すこと(
    mock_abspath, mock_isdir
):
    assert resolve_search_directory("docs") == r"C:\docs"


@patch("main.os.path.isdir", return_value=True)
@patch("main.os.path.abspath", return_value=r"C:\docs")
def test正常系_resolve_search_directoryメソッドが前後の引用符を除去すること(
    mock_abspath, mock_isdir
):
    resolve_search_directory(r'"C:\docs"')
    mock_isdir.assert_called_once_with(r"C:\docs")


def test異常系_resolve_search_directoryメソッドが空文字でSearchDirectoryErrorを出すこと():
    with pytest.raises(SearchDirectoryError, match="検索対象フォルダを指定してください。"):
        resolve_search_directory("")


def test異常系_resolve_search_directoryメソッドが空白のみでSearchDirectoryErrorを出すこと():
    with pytest.raises(SearchDirectoryError, match="検索対象フォルダを指定してください。"):
        resolve_search_directory("   ")


@patch("main.os.path.isdir", return_value=False)
def test異常系_resolve_search_directoryメソッドが存在しないパスでSearchDirectoryErrorを出すこと(
    mock_isdir,
):
    with pytest.raises(SearchDirectoryError, match="指定されたフォルダが存在しません。"):
        resolve_search_directory(r"C:\missing")


def test正常系___init__メソッドがsearch_stringを保持すること(search_thread):
    assert search_thread.search_string == "needle"


def test正常系___init__メソッドがsearch_directoryを保持すること(search_thread):
    assert search_thread.search_directory == r"C:\docs"


def test正常系___init__メソッドがfiles_searchedを0にすること(search_thread):
    assert search_thread.files_searched == 0


def test正常系___init__メソッドがfiles_excludedを0にすること(search_thread):
    assert search_thread.files_excluded == 0


def test正常系___init__メソッドがexcluded_filesを空リストにすること(search_thread):
    assert search_thread.excluded_files == []


def test正常系__exclude_fileメソッドが除外件数を1増やすこと(search_thread):
    search_thread._exclude_file(r"C:\docs\bad.txt")
    assert search_thread.files_excluded == 1


def test正常系__exclude_fileメソッドがパスをexcluded_filesへ入れること(search_thread):
    search_thread._exclude_file(r"C:\docs\bad.txt")
    assert search_thread.excluded_files == [r"C:\docs\bad.txt"]


def test正常系_runメソッドがsearch_filesを指定ディレクトリで呼ぶこと(search_thread):
    search_thread.search_files = MagicMock()
    search_thread.search_completed = MagicMock()
    search_thread.run()
    search_thread.search_files.assert_called_once_with("needle", r"C:\docs")


def test正常系_runメソッドがsearch_completedを件数付きで通知すること(search_thread):
    search_thread.search_files = MagicMock()
    search_thread.search_completed = MagicMock()
    search_thread.files_searched = 3
    search_thread.files_excluded = 1
    search_thread.excluded_files = [r"C:\docs\bad.txt"]
    search_thread.run()
    search_thread.search_completed.emit.assert_called_once_with(
        3, 1, [r"C:\docs\bad.txt"]
    )


@patch("main.os.walk", return_value=[])
def test正常系_search_filesメソッドがファイル0件のとき進捗100を通知すること(
    mock_walk, search_thread
):
    search_thread.progress_update = MagicMock()
    search_thread.search_files("needle", r"C:\empty")
    search_thread.progress_update.emit.assert_called_once_with(100)


@patch("main.chardet.detect", return_value={"encoding": "utf-8"})
@patch("main.os.walk")
def test正常系_search_filesメソッドが一致するtxtをresult_foundで通知すること(
    mock_walk, mock_detect, search_thread
):
    hit_path = os.path.join(r"C:\docs", "hit.txt")
    mock_walk.return_value = [(r"C:\docs", [], ["hit.txt"])]
    search_thread.result_found = MagicMock()
    mocked_file = MagicMock()
    mocked_file.read.return_value = b"needle here"
    mocked_file.__enter__.return_value = mocked_file
    with patch("builtins.open", return_value=mocked_file):
        search_thread.search_files("needle", r"C:\docs")
    search_thread.result_found.emit.assert_called_once_with(hit_path)


@patch("main.chardet.detect", return_value={"encoding": "utf-8"})
@patch("main.os.walk")
def test正常系_search_filesメソッドが不一致のtxtをresult_foundしないこと(
    mock_walk, mock_detect, search_thread
):
    mock_walk.return_value = [(r"C:\docs", [], ["miss.txt"])]
    search_thread.result_found = MagicMock()
    mocked_file = MagicMock()
    mocked_file.read.return_value = b"nothing"
    mocked_file.__enter__.return_value = mocked_file
    with patch("builtins.open", return_value=mocked_file):
        search_thread.search_files("needle", r"C:\docs")
    search_thread.result_found.emit.assert_not_called()


@patch("main.docx.Document")
@patch("main.os.walk")
def test正常系_search_filesメソッドが一致するdocxをresult_foundで通知すること(
    mock_walk, mock_document, search_thread
):
    hit_path = os.path.join(r"C:\docs", "hit.docx")
    mock_walk.return_value = [(r"C:\docs", [], ["hit.docx"])]
    paragraph = MagicMock()
    paragraph.text = "needle"
    mock_document.return_value.paragraphs = [paragraph]
    search_thread.result_found = MagicMock()
    search_thread.search_files("needle", r"C:\docs")
    search_thread.result_found.emit.assert_called_once_with(hit_path)


@patch("main.os.walk")
def test正常系_search_filesメソッドが開いているdocxを除外すること(mock_walk, search_thread):
    excluded_path = os.path.join(r"C:\docs", "~$open.docx")
    mock_walk.return_value = [(r"C:\docs", [], ["~$open.docx"])]
    search_thread.search_files("needle", r"C:\docs")
    assert search_thread.excluded_files == [excluded_path]


@patch("main.chardet.detect", return_value={"encoding": None})
@patch("main.os.walk")
def test異常系_search_filesメソッドがencodingがNoneのとき除外すること(
    mock_walk, mock_detect, search_thread
):
    excluded_path = os.path.join(r"C:\docs", "unknown.txt")
    mock_walk.return_value = [(r"C:\docs", [], ["unknown.txt"])]
    mocked_file = MagicMock()
    mocked_file.read.return_value = b"\xff\xfe"
    mocked_file.__enter__.return_value = mocked_file
    with patch("builtins.open", return_value=mocked_file):
        search_thread.search_files("needle", r"C:\docs")
    assert search_thread.excluded_files == [excluded_path]


@patch("main.os.path.getsize", return_value=MAX_SEARCH_FILE_BYTES + 1)
@patch("main.os.walk")
def test異常系_search_filesメソッドが過大ファイルを除外すること(
    mock_walk, mock_getsize, search_thread
):
    excluded_path = os.path.join(r"C:\docs", "huge.txt")
    mock_walk.return_value = [(r"C:\docs", [], ["huge.txt"])]
    search_thread.search_files("needle", r"C:\docs")
    assert search_thread.excluded_files == [excluded_path]


@patch("main.docx.Document", side_effect=RuntimeError("broken"))
@patch("main.os.walk")
def test異常系_search_filesメソッドが予期しない例外で除外すること(
    mock_walk, mock_document, search_thread
):
    excluded_path = os.path.join(r"C:\docs", "boom.docx")
    mock_walk.return_value = [(r"C:\docs", [], ["boom.docx"])]
    search_thread.search_files("needle", r"C:\docs")
    assert search_thread.excluded_files == [excluded_path]


@patch("main.chardet.detect", return_value={"encoding": "utf-8"})
@patch("main.os.walk")
def test異常系_search_filesメソッドがUnicodeDecodeErrorで除外すること(
    mock_walk, mock_detect, search_thread
):
    excluded_path = os.path.join(r"C:\docs", "bad.txt")
    mock_walk.return_value = [(r"C:\docs", [], ["bad.txt"])]
    mocked_file = MagicMock()
    mocked_bytes = MagicMock()
    mocked_bytes.decode.side_effect = UnicodeDecodeError("utf-8", b"x", 0, 1, "reason")
    mocked_file.read.return_value = mocked_bytes
    mocked_file.__enter__.return_value = mocked_file
    with patch("builtins.open", return_value=mocked_file):
        search_thread.search_files("needle", r"C:\docs")
    assert search_thread.excluded_files == [excluded_path]


@patch("main.docx.Document")
@patch("main.os.walk")
def test異常系_search_filesメソッドがPackageNotFoundErrorで除外すること(
    mock_walk, mock_document, search_thread
):
    import docx as docx_module

    excluded_path = os.path.join(r"C:\docs", "broken.docx")
    mock_walk.return_value = [(r"C:\docs", [], ["broken.docx"])]
    mock_document.side_effect = docx_module.opc.exceptions.PackageNotFoundError("broken")
    search_thread.search_files("needle", r"C:\docs")
    assert search_thread.excluded_files == [excluded_path]


def test正常系___init__メソッドがフォルダ入力欄を検索欄の上に置くこと(window):
    window.show()
    folder_y = window.folder_input.mapTo(window, window.folder_input.rect().topLeft()).y()
    search_y = window.search_input.mapTo(window, window.search_input.rect().topLeft()).y()
    assert folder_y < search_y


def test正常系___init__メソッドが検索対象フォルダのプレースホルダを設定すること(window):
    assert window.folder_input.placeholderText() == "検索対象フォルダのパスを入力してください"


@patch("main.QFileDialog.getExistingDirectory", return_value=r"C:/docs")
def test正常系_browse_directoryメソッドが選択フォルダを入力欄へ入れること(mock_dialog, window):
    window.browse_directory()
    assert window.folder_input.text() == os.path.normpath(r"C:/docs")


@patch("main.QFileDialog.getExistingDirectory", return_value="")
def test正常系_browse_directoryメソッドがキャンセル時に入力欄を変えないこと(mock_dialog, window):
    window.folder_input.setText(r"C:\keep")
    window.browse_directory()
    assert window.folder_input.text() == r"C:\keep"


def test異常系_start_searchメソッドが検索文字列未入力でステータスを出すこと(window):
    window.search_input.setText("")
    window.start_search()
    assert window.status_bar.currentMessage() == "検索文字列を入力してください。"


@patch("main.resolve_search_directory", side_effect=SearchDirectoryError("指定されたフォルダが存在しません。"))
def test異常系_start_searchメソッドが不正フォルダでステータスを出すこと(mock_resolve, window):
    window.search_input.setText("needle")
    window.folder_input.setText(r"C:\missing")
    window.start_search()
    assert window.status_bar.currentMessage() == "指定されたフォルダが存在しません。"


@patch("main.SearchThread")
@patch("main.resolve_search_directory", return_value=r"C:\docs")
def test正常系_start_searchメソッドが指定フォルダでSearchThreadを作ること(
    mock_resolve, mock_thread_cls, window
):
    mock_thread_cls.return_value = MagicMock()
    window.search_input.setText("needle")
    window.folder_input.setText(r"C:\docs")
    window.start_search()
    mock_thread_cls.assert_called_once_with("needle", r"C:\docs")


@patch("main.SearchThread")
@patch("main.resolve_search_directory", return_value=r"C:\docs")
def test正常系_start_searchメソッドがスレッドを開始すること(
    mock_resolve, mock_thread_cls, window
):
    mock_thread = MagicMock()
    mock_thread_cls.return_value = mock_thread
    window.search_input.setText("needle")
    window.folder_input.setText(r"C:\docs")
    window.start_search()
    mock_thread.start.assert_called_once()


@patch("main.SearchThread")
@patch("main.resolve_search_directory", return_value=r"C:\docs")
def test正常系_start_searchメソッドが検索中に検索ボタンを無効化すること(
    mock_resolve, mock_thread_cls, window
):
    mock_thread_cls.return_value = MagicMock()
    window.search_input.setText("needle")
    window.folder_input.setText(r"C:\docs")
    window.start_search()
    assert window.search_button.isEnabled() is False


def test正常系_update_progressメソッドがプログレスバーへ値を入れること(window):
    window.update_progress(40)
    assert window.progress_bar.value() == 40


def test正常系_update_progressメソッドがステータスへ進捗を出すこと(window):
    window.update_progress(40)
    assert window.status_bar.currentMessage() == "検索進捗: 40%"


def test正常系_display_resultメソッドが結果エリアへパスを追加すること(window):
    window.display_result(r"C:\docs\hit.txt")
    assert r"C:\docs\hit.txt" in window.result_area.toPlainText()


def test正常系_search_completedメソッドが検索ボタンを再有効化すること(window):
    window.search_button.setEnabled(False)
    window.search_completed(1, 0, [])
    assert window.search_button.isEnabled() is True


def test正常系_search_completedメソッドがフォルダ入力を再有効化すること(window):
    window.folder_input.setEnabled(False)
    window.search_completed(1, 0, [])
    assert window.folder_input.isEnabled() is True


def test正常系_search_completedメソッドが除外ファイルを結果エリアへ出すこと(window):
    window.search_completed(2, 1, [r"C:\docs\bad.txt"])
    assert r"C:\docs\bad.txt" in window.result_area.toPlainText()
