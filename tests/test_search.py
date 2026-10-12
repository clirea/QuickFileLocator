from pathlib import Path
from unittest.mock import MagicMock, call, patch

import pytest

from quickfilelocator.query import SearchQuery
from quickfilelocator.readers import UnreadableFileError
from quickfilelocator.search import ExcludedFile, FileSearcher, SearchDirectoryError


@pytest.fixture
def reader():
    """ファイルの中身をそのまま本文として返す読み取り手段。"""
    mock = MagicMock()
    mock.read.side_effect = lambda path: path.read_text(encoding="utf-8")
    return mock


@pytest.fixture
def searcher(reader):
    return FileSearcher({".txt": reader})


@pytest.fixture
def tree(tmp_path):
    """一致する2件・一致しない1件・対象外1件を持つフォルダー。"""
    (tmp_path / "sub").mkdir()
    (tmp_path / "b.txt").write_text("needle", encoding="utf-8")
    (tmp_path / "a.txt").write_text("hay", encoding="utf-8")
    (tmp_path / "sub" / "c.TXT").write_text("needle in sub", encoding="utf-8")
    (tmp_path / "image.png").write_text("needle", encoding="utf-8")
    return tmp_path


def run_search(searcher, directory, text="needle", **overrides):
    callbacks = {"on_match": MagicMock(), "on_progress": MagicMock(), **overrides}
    summary = searcher.search(SearchQuery.parse(text), directory, **callbacks)
    return summary, callbacks


def test正常系_searchメソッドが一致したファイルを名前順に通知すること(searcher, tree):
    _, callbacks = run_search(searcher, tree)

    assert callbacks["on_match"].call_args_list == [call(tree / "b.txt"), call(tree / "sub" / "c.TXT")]


def test正常系_searchメソッドが対象拡張子のファイルだけを対象件数に数えること(searcher, tree):
    summary, _ = run_search(searcher, tree)

    assert summary.target_count == 3


def test正常系_searchメソッドが照合できた件数を返すこと(searcher, tree):
    summary, _ = run_search(searcher, tree)

    assert summary.searched_count == 3


def test正常系_searchメソッドが一致した件数を返すこと(searcher, tree):
    summary, _ = run_search(searcher, tree)

    assert summary.matched_count == 2


def test正常系_searchメソッドが最後まで検索したときキャンセルなしとして返すこと(searcher, tree):
    summary, _ = run_search(searcher, tree)

    assert summary.cancelled is False


def test正常系_searchメソッドが対象件数の確定時と1件ごとに進捗を通知すること(searcher, tree):
    _, callbacks = run_search(searcher, tree)

    assert callbacks["on_progress"].call_args_list == [call(0, 3), call(1, 3), call(2, 3), call(3, 3)]


def test正常系_searchメソッドが対象ファイルのないフォルダーで件数0を返すこと(searcher, tmp_path):
    summary, _ = run_search(searcher, tmp_path)

    assert summary.target_count == 0


def test正常系_searchメソッドが対象ファイルのないフォルダーでも進捗を通知すること(searcher, tmp_path):
    _, callbacks = run_search(searcher, tmp_path)

    callbacks["on_progress"].assert_called_once_with(0, 0)


def test正常系_searchメソッドが読み取れないファイルを理由つきで除外すること(searcher, reader, tree):
    reader.read.side_effect = UnreadableFileError("文字コードを判定できません")

    summary, _ = run_search(searcher, tree)

    assert summary.excluded[0] == ExcludedFile(tree / "a.txt", "文字コードを判定できません")


def test正常系_searchメソッドが読み取れないファイルを照合件数に数えないこと(searcher, reader, tree):
    reader.read.side_effect = UnreadableFileError("文字コードを判定できません")

    summary, _ = run_search(searcher, tree)

    assert summary.searched_count == 0


def test正常系_searchメソッドがOSエラーのファイルを除外して残りの検索を続けること(searcher, reader, tree):
    reader.read.side_effect = [PermissionError(13, "Permission denied"), "needle", "needle"]

    summary, _ = run_search(searcher, tree)

    assert summary.matched_count == 2


def test正常系_searchメソッドがOSエラーの内容を除外理由に含めること(searcher, reader, tree):
    reader.read.side_effect = [PermissionError(13, "Permission denied"), "needle", "needle"]

    summary, _ = run_search(searcher, tree)

    assert summary.excluded == (ExcludedFile(tree / "a.txt", "読み取れません (Permission denied)"),)


def test正常系_searchメソッドが除外したファイルも進捗に数えること(searcher, reader, tree):
    reader.read.side_effect = UnreadableFileError("x")

    _, callbacks = run_search(searcher, tree)

    assert callbacks["on_progress"].call_args == call(3, 3)


def test正常系_searchメソッドが上限を超えるサイズのファイルを読まずに除外すること(reader, tmp_path):
    (tmp_path / "big.txt").write_text("needle" * 10, encoding="utf-8")

    run_search(FileSearcher({".txt": reader}, max_file_bytes=10), tmp_path)

    reader.read.assert_not_called()


def test正常系_searchメソッドが上限ちょうどのサイズのファイルを検索すること(reader, tmp_path):
    (tmp_path / "edge.txt").write_text("needle", encoding="utf-8")

    summary, _ = run_search(FileSearcher({".txt": reader}, max_file_bytes=6), tmp_path)

    assert summary.matched_count == 1


def test正常系_searchメソッドがサイズ上限を除外理由にMB単位で示すこと(reader, tmp_path):
    (tmp_path / "big.txt").write_bytes(b"x" * (1024 * 1024 + 1))

    summary, _ = run_search(FileSearcher({".txt": reader}, max_file_bytes=1024 * 1024), tmp_path)

    assert summary.excluded == (ExcludedFile(tmp_path / "big.txt", "サイズが上限(1MB)を超えています"),)


def test正常系_searchメソッドが読み取れないフォルダーを除外に記録すること(searcher, tmp_path):
    def fake_walk(directory, onerror):
        onerror(PermissionError(13, "Permission denied", str(tmp_path / "secret")))
        return iter([])

    with patch("quickfilelocator.search.os.walk", side_effect=fake_walk):
        summary, _ = run_search(searcher, tmp_path)

    assert summary.excluded == (ExcludedFile(tmp_path / "secret", "読み取れません (Permission denied)"),)


def test正常系_searchメソッドがサブフォルダーを名前順にたどるよう走査へ指示すること(searcher, tmp_path):
    dir_names = ["z", "a", "m"]

    with patch("quickfilelocator.search.os.walk", return_value=iter([(str(tmp_path), dir_names, [])])):
        run_search(searcher, tmp_path)

    assert dir_names == ["a", "m", "z"]


def test正常系_searchメソッドが大文字で登録された拡張子も対象にすること(reader, tree):
    summary, _ = run_search(FileSearcher({".TXT": reader}), tree)

    assert summary.target_count == 3


def test正常系_searchメソッドが開始前のキャンセルで何も読まないこと(searcher, reader, tree):
    run_search(searcher, tree, is_cancelled=lambda: True)

    reader.read.assert_not_called()


def test正常系_searchメソッドが開始前のキャンセルをキャンセルありとして返すこと(searcher, tree):
    summary, _ = run_search(searcher, tree, is_cancelled=lambda: True)

    assert summary.cancelled is True


@pytest.fixture
def cancel_after_first_read(reader):
    """1件目を読み終えた時点でキャンセルが要求される状況を作る。"""
    state = {"cancelled": False}

    def read(path):
        state["cancelled"] = True
        return path.read_text(encoding="utf-8")

    reader.read.side_effect = read
    return lambda: state["cancelled"]


def test正常系_searchメソッドが途中のキャンセルで以降のファイルを読まないこと(
    searcher, reader, tree, cancel_after_first_read
):
    run_search(searcher, tree, is_cancelled=cancel_after_first_read)

    assert reader.read.call_count == 1


def test正常系_searchメソッドが途中のキャンセルまでの結果を返すこと(searcher, tree, cancel_after_first_read):
    summary, _ = run_search(searcher, tree, is_cancelled=cancel_after_first_read)

    assert (summary.searched_count, summary.cancelled) == (1, True)


def test正常系_searchメソッドが走査中のキャンセルで対象を集めるのをやめること(searcher, tree):
    summary, _ = run_search(searcher, tree, is_cancelled=MagicMock(side_effect=[False, True, True]))

    assert summary.target_count == 2


@pytest.mark.parametrize("name", ["missing", "b.txt"])
def test異常系_searchメソッドがフォルダーでない検索先でSearchDirectoryErrorを送出すること(searcher, tree, name):
    with pytest.raises(SearchDirectoryError, match="検索先フォルダー"):
        run_search(searcher, Path(tree, name))
