import threading
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest
from PyQt5.QtGui import QCloseEvent
from PyQt5.QtWidgets import QApplication

from quickfilelocator import gui
from quickfilelocator.gui import MainWindow, SearchThread
from quickfilelocator.query import SearchQuery
from quickfilelocator.readers import default_readers
from quickfilelocator.search import ExcludedFile, FileSearcher, SearchDirectoryError, SearchSummary

QUERY = SearchQuery.parse("needle")


def make_summary(**overrides):
    values = {"target_count": 5, "searched_count": 4, "matched_count": 2, "excluded": (), "cancelled": False}
    return SearchSummary(**{**values, **overrides})


@pytest.fixture
def searcher():
    mock = MagicMock()
    mock.search.return_value = make_summary()
    return mock


@pytest.fixture
def thread(qapp, searcher):
    return SearchThread(searcher, QUERY, Path("/data"))


def collect(signal):
    received = []
    signal.connect(lambda *args: received.append(args))
    return received


def test正常系_runメソッドが検索条件と検索先を検索処理へ渡すこと(thread, searcher):
    thread.run()

    assert searcher.search.call_args.args == (QUERY, Path("/data"))


def test正常系_runメソッドが検索結果を完了シグナルで通知すること(thread, searcher):
    completed = collect(thread.search_completed)

    thread.run()

    assert completed == [(searcher.search.return_value,)]


def test正常系_runメソッドが一致したパスを文字列で通知すること(thread, searcher):
    searcher.search.side_effect = lambda *args, on_match, **kwargs: on_match(Path("/data/a.txt"))
    found = collect(thread.result_found)

    thread.run()

    assert found == [(str(Path("/data/a.txt")),)]


def test正常系_runメソッドが中断要求の有無をキャンセル判定として渡すこと(thread, searcher):
    thread.run()

    assert searcher.search.call_args.kwargs["is_cancelled"] == thread.isInterruptionRequested


def test異常系_runメソッドが検索先のエラーを失敗シグナルで通知すること(thread, searcher):
    searcher.search.side_effect = SearchDirectoryError("検索先フォルダーが見つかりません: /data")
    failed = collect(thread.search_failed)

    thread.run()

    assert failed == [("検索先フォルダーが見つかりません: /data",)]


def test異常系_runメソッドが想定外の例外を送出せず失敗シグナルで通知すること(thread, searcher):
    searcher.search.side_effect = RuntimeError("boom")
    failed = collect(thread.search_failed)

    thread.run()

    assert failed == [("検索中にエラーが発生しました: RuntimeError: boom",)]


def test異常系_runメソッドが失敗時に完了シグナルを出さないこと(thread, searcher):
    searcher.search.side_effect = RuntimeError("boom")
    completed = collect(thread.search_completed)

    thread.run()

    assert completed == []


def test正常系_runメソッドが進捗を1パーセント刻みに間引いて通知すること(thread, searcher):
    def search(*args, on_progress, **kwargs):
        for processed_count in range(1001):
            on_progress(processed_count, 1000)

    searcher.search.side_effect = search
    progress = collect(thread.progress_update)

    thread.run()

    assert len(progress) == 101


def test正常系_runメソッドが最後の進捗を必ず通知すること(thread, searcher):
    searcher.search.side_effect = lambda *args, on_progress, **kwargs: [on_progress(n, 3) for n in range(4)]
    progress = collect(thread.progress_update)

    thread.run()

    assert progress[-1] == (3, 3)


def test正常系_runメソッドが対象0件の進捗を通知すること(thread, searcher):
    searcher.search.side_effect = lambda *args, on_progress, **kwargs: on_progress(0, 0)
    progress = collect(thread.progress_update)

    thread.run()

    assert progress == [(0, 0)]


@pytest.fixture
def thread_factory():
    return MagicMock()


@pytest.fixture
def open_path():
    return MagicMock(return_value=True)


@pytest.fixture
def window(qapp, searcher, thread_factory, open_path, tmp_path):
    main_window = MainWindow(searcher, tmp_path, thread_factory=thread_factory, open_path=open_path)
    main_window.search_input.setText("needle")
    yield main_window
    main_window.search_thread = None
    main_window.close()


@pytest.fixture
def searching_window(window):
    window.start_search()
    return window


def test正常系_initメソッドが初期の検索先を入力欄に表示すること(window, tmp_path):
    assert window.directory_input.text() == str(tmp_path)


def test正常系_initメソッドが大文字小文字を区別する設定で開始すること(window):
    assert window.case_checkbox.isChecked() is True


def test正常系_initメソッドがキャンセルボタンを無効にしておくこと(window):
    assert window.cancel_button.isEnabled() is False


def test正常系_choose_directoryメソッドが選択されたフォルダーを入力欄に反映すること(window):
    with patch.object(gui.QFileDialog, "getExistingDirectory", return_value="/picked"):
        window.choose_directory()

    assert window.directory_input.text() == "/picked"


def test正常系_choose_directoryメソッドが選択の取り消しで入力欄を変えないこと(window, tmp_path):
    with patch.object(gui.QFileDialog, "getExistingDirectory", return_value=""):
        window.choose_directory()

    assert window.directory_input.text() == str(tmp_path)


def test正常系_start_searchメソッドが解析した検索条件と検索先でスレッドを作ること(
    searching_window, searcher, thread_factory, tmp_path
):
    thread_factory.assert_called_once_with(searcher, QUERY, tmp_path)


def test正常系_start_searchメソッドが大文字小文字の設定を検索条件に反映すること(window, thread_factory):
    window.case_checkbox.setChecked(False)

    window.start_search()

    assert thread_factory.call_args.args[1].case_sensitive is False


def test正常系_start_searchメソッドがスレッドを開始すること(searching_window, thread_factory):
    thread_factory.return_value.start.assert_called_once_with()


def test正常系_start_searchメソッドが終了シグナルを後片付けへ接続すること(searching_window, thread_factory):
    thread_factory.return_value.finished.connect.assert_called_once_with(searching_window.search_finished)


def test正常系_start_searchメソッドが検索中は検索ボタンを無効にすること(searching_window):
    assert searching_window.search_button.isEnabled() is False


def test正常系_start_searchメソッドが検索中はキャンセルボタンを有効にすること(searching_window):
    assert searching_window.cancel_button.isEnabled() is True


def test正常系_start_searchメソッドが前回の結果を消すこと(window):
    window.display_result("/old/result.txt")

    window.start_search()

    assert window.result_list.count() == 0


def test正常系_start_searchメソッドが前回の除外一覧を隠すこと(window):
    window.search_completed(make_summary(excluded=(ExcludedFile(Path("/old.txt"), "理由"),)))

    window.start_search()

    assert window.excluded_area.isHidden() is True


def test正常系_start_searchメソッドが件数確定まで進捗バーを処理中表示にすること(searching_window):
    assert searching_window.progress_bar.maximum() == 0


def test正常系_start_searchメソッドが検索中の再実行でスレッドを増やさないこと(searching_window, thread_factory):
    searching_window.start_search()

    assert thread_factory.call_count == 1


@pytest.mark.parametrize("text", ["", "   ", '"閉じていない'])
def test異常系_start_searchメソッドが解釈できない検索式でスレッドを作らないこと(window, thread_factory, text):
    window.search_input.setText(text)

    window.start_search()

    thread_factory.assert_not_called()


def test異常系_start_searchメソッドが解釈できない検索式の理由を表示すること(window):
    window.search_input.setText('"閉じていない')

    window.start_search()

    assert window.status_bar.currentMessage() == '引用符(")が閉じられていません。'


@pytest.mark.parametrize("directory", ["", "/no/such/folder/for/test"])
def test異常系_start_searchメソッドが存在しない検索先でスレッドを作らないこと(window, thread_factory, directory):
    window.directory_input.setText(directory)

    window.start_search()

    thread_factory.assert_not_called()


def test異常系_start_searchメソッドが存在しない検索先を表示すること(window):
    window.directory_input.setText("/no/such/folder/for/test")

    window.start_search()

    assert window.status_bar.currentMessage() == "検索先フォルダーが見つかりません: /no/such/folder/for/test"


def test異常系_start_searchメソッドが入力エラーのあとも検索ボタンを有効に保つこと(window):
    window.search_input.setText("")

    window.start_search()

    assert window.search_button.isEnabled() is True


def test正常系_cancel_searchメソッドがスレッドへ中断を要求すること(searching_window, thread_factory):
    searching_window.cancel_search()

    thread_factory.return_value.requestInterruption.assert_called_once_with()


def test正常系_cancel_searchメソッドがキャンセルボタンを無効にすること(searching_window):
    searching_window.cancel_search()

    assert searching_window.cancel_button.isEnabled() is False


def test正常系_cancel_searchメソッドが検索していないとき表示を変えないこと(window):
    window.cancel_search()

    assert window.status_bar.currentMessage() == ""


def test正常系_update_progressメソッドが進捗バーの最大値を対象件数にすること(window):
    window.update_progress(3, 12)

    assert window.progress_bar.maximum() == 12


def test正常系_update_progressメソッドが進捗バーの値を処理済み件数にすること(window):
    window.update_progress(3, 12)

    assert window.progress_bar.value() == 3


def test正常系_update_progressメソッドが件数を表示すること(window):
    window.update_progress(3, 12)

    assert window.status_bar.currentMessage() == "検索中: 3 / 12 件"


def test正常系_update_progressメソッドが対象0件でも進捗バーを処理中表示に戻さないこと(window):
    window.update_progress(0, 0)

    assert window.progress_bar.maximum() == 1


def test正常系_display_resultメソッドが結果を一覧に追加すること(window):
    window.display_result("/data/a.txt")

    assert window.result_list.item(0).text() == "/data/a.txt"


def test正常系_search_completedメソッドが件数の内訳を表示すること(window):
    window.search_completed(make_summary())

    assert window.status_bar.currentMessage() == "検索完了。 対象 5 件中 4 件を検索し、2 件が一致、0 件を除外しました。"


def test正常系_search_completedメソッドがキャンセルされた結果であることを表示すること(window):
    window.search_completed(make_summary(cancelled=True))

    assert window.status_bar.currentMessage().startswith("検索をキャンセルしました。")


def test正常系_search_completedメソッドが除外したファイルを理由つきで表示すること(window):
    window.search_completed(make_summary(excluded=(ExcludedFile(Path("/data/x.txt"), "文字コードを判定できません"),)))

    assert window.excluded_area.toPlainText() == f"{Path('/data/x.txt')}  [文字コードを判定できません]"


def test正常系_search_completedメソッドが除外があるとき除外一覧を表示状態にすること(window):
    window.search_completed(make_summary(excluded=(ExcludedFile(Path("/data/x.txt"), "理由"),)))

    assert window.excluded_area.isHidden() is False


def test正常系_search_completedメソッドが除外がないとき除外一覧を隠したままにすること(window):
    window.search_completed(make_summary())

    assert window.excluded_area.isHidden() is True


def test正常系_search_failedメソッドが失敗の内容を表示すること(window):
    window.search_failed("検索中にエラーが発生しました: RuntimeError: boom")

    assert window.status_bar.currentMessage() == "検索中にエラーが発生しました: RuntimeError: boom"


def test正常系_search_finishedメソッドが検索ボタンを有効に戻すこと(searching_window):
    searching_window.search_finished()

    assert searching_window.search_button.isEnabled() is True


def test正常系_search_finishedメソッドがキャンセルボタンを無効に戻すこと(searching_window):
    searching_window.search_finished()

    assert searching_window.cancel_button.isEnabled() is False


def test正常系_search_finishedメソッドが進捗バーを隠すこと(searching_window):
    searching_window.search_finished()

    assert searching_window.progress_bar.isHidden() is True


def test正常系_search_finishedメソッドがスレッドの破棄を予約すること(searching_window, thread_factory):
    searching_window.search_finished()

    thread_factory.return_value.deleteLater.assert_called_once_with()


def test正常系_search_finishedメソッドが次の検索を開始できる状態に戻すこと(searching_window, thread_factory):
    searching_window.search_finished()

    searching_window.start_search()

    assert thread_factory.call_count == 2


def test正常系_search_finishedメソッドが終了要求を受けていたとき画面を閉じること(searching_window):
    searching_window.closeEvent(QCloseEvent())

    with patch.object(searching_window, "close") as close:
        searching_window.search_finished()

    close.assert_called_once_with()


def test正常系_search_finishedメソッドが終了要求がないとき画面を閉じないこと(searching_window):
    with patch.object(searching_window, "close") as close:
        searching_window.search_finished()

    close.assert_not_called()


def test正常系_open_resultメソッドが選ばれた結果のパスを開くこと(window, open_path):
    window.display_result("/data/a.txt")

    window.open_result(window.result_list.item(0))

    open_path.assert_called_once_with("/data/a.txt")


def test異常系_open_resultメソッドが開けなかったことを表示すること(window, open_path):
    open_path.return_value = False
    window.display_result("/data/a.txt")

    window.open_result(window.result_list.item(0))

    assert window.status_bar.currentMessage() == "ファイルを開けませんでした: /data/a.txt"


def test正常系_copy_selected_resultsメソッドが選択したパスを改行区切りでコピーすること(window, qapp):
    window.display_result("/data/a.txt")
    window.display_result("/data/b.txt")
    window.result_list.selectAll()

    window.copy_selected_results()

    assert qapp.clipboard().text() == "/data/a.txt\n/data/b.txt"


def test正常系_copy_selected_resultsメソッドが未選択のときクリップボードを変えないこと(window, qapp):
    qapp.clipboard().setText("keep")
    window.display_result("/data/a.txt")

    window.copy_selected_results()

    assert qapp.clipboard().text() == "keep"


def test正常系_closeEventメソッドが検索していないとき終了を受け入れること(window):
    event = QCloseEvent()

    window.closeEvent(event)

    assert event.isAccepted() is True


def test正常系_closeEventメソッドが検索中は終了を保留すること(searching_window):
    event = QCloseEvent()

    searching_window.closeEvent(event)

    assert event.isAccepted() is False


def test正常系_closeEventメソッドが検索中はスレッドへ中断を要求すること(searching_window, thread_factory):
    searching_window.closeEvent(QCloseEvent())

    thread_factory.return_value.requestInterruption.assert_called_once_with()


def test正常系_open_with_default_appメソッドがローカルファイルのURLをOSへ渡すこと():
    with patch.object(gui.QDesktopServices, "openUrl", return_value=True) as open_url:
        gui.open_with_default_app("/data/a.txt")

    assert open_url.call_args.args[0].toLocalFile() == "/data/a.txt"


def wait_until_idle(window, timeout_ms=5000):
    """実スレッドの終了と、画面側へ届いたシグナルの処理を待つ。"""
    thread = window.search_thread
    if thread is not None:
        assert thread.wait(timeout_ms)
    QApplication.processEvents()


@pytest.fixture
def real_window(qapp, tmp_path):
    """実際の読み取り・走査・スレッドをつないだ画面。"""
    (tmp_path / "hit.txt").write_text("needle", encoding="utf-8")
    (tmp_path / "broken.docx").write_bytes(b"not a docx")
    main_window = MainWindow(FileSearcher(default_readers()), tmp_path)
    main_window.search_input.setText("needle")
    yield main_window
    wait_until_idle(main_window)
    main_window.close()


def test正常系_start_searchメソッドが実スレッドの検索結果を一覧に表示すること(real_window, tmp_path):
    real_window.start_search()
    wait_until_idle(real_window)

    assert [real_window.result_list.item(i).text() for i in range(real_window.result_list.count())] == [
        str(tmp_path / "hit.txt")
    ]


def test正常系_start_searchメソッドが実スレッドの終了後に検索ボタンを有効に戻すこと(real_window):
    real_window.start_search()
    wait_until_idle(real_window)

    assert real_window.search_button.isEnabled() is True


def test正常系_start_searchメソッドが壊れたファイルを除外一覧に出して完了すること(real_window, tmp_path):
    real_window.start_search()
    wait_until_idle(real_window)

    assert real_window.excluded_area.toPlainText().startswith(str(tmp_path / "broken.docx"))


def test正常系_closeEventメソッドが実スレッドの検索中に閉じても停止を待って終了すること(qapp, tmp_path):
    (tmp_path / "a.txt").write_text("needle", encoding="utf-8")
    reading = threading.Event()
    release = threading.Event()
    reader = MagicMock()
    reader.read.side_effect = lambda path: (reading.set(), release.wait(5), "needle")[2]
    main_window = MainWindow(FileSearcher({".txt": reader}), tmp_path)
    main_window.search_input.setText("needle")
    main_window.show()
    main_window.start_search()
    reading.wait(5)

    main_window.close()
    still_open_while_running = main_window.isVisible()
    release.set()
    wait_until_idle(main_window)

    assert (still_open_while_running, main_window.isVisible()) == (True, False)
