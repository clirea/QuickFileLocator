from quickfilelocator.settings import (
    as_bool,
    as_string_list,
    existing_directory,
    existing_recent_folders,
    push_recent_folder,
)


def test正常系_as_boolメソッドが文字列のtrueを真にすること():
    assert as_bool("true", False) is True


def test正常系_as_boolメソッドが不明な値で既定を返すこと():
    assert as_bool("maybe", True) is True


def test正常系_as_string_listメソッドが単一文字列をリストにすること():
    assert as_string_list("C:\\docs") == ["C:\\docs"]


def test正常系_existing_directoryメソッドが存在するフォルダーを返すこと(tmp_path):
    assert existing_directory(tmp_path) == tmp_path


def test異常系_existing_directoryメソッドが無いパスでNoneを返すこと(tmp_path):
    assert existing_directory(tmp_path / "missing") is None


def test正常系_push_recent_folderメソッドが先頭に追加すること(tmp_path):
    folder = tmp_path / "docs"
    folder.mkdir()

    assert push_recent_folder([], str(folder)) == [str(folder)]


def test正常系_push_recent_folderメソッドが重複を先頭へ移すこと(tmp_path):
    first = tmp_path / "a"
    second = tmp_path / "b"
    first.mkdir()
    second.mkdir()
    folders = push_recent_folder(push_recent_folder([], str(first)), str(second))

    assert push_recent_folder(folders, str(first)) == [str(first), str(second)]


def test正常系_push_recent_folderメソッドが11件目で古いフォルダーを捨てること(tmp_path):
    folders: list[str] = []
    created: list[str] = []
    for index in range(11):
        folder = tmp_path / f"d{index}"
        folder.mkdir()
        created.append(str(folder))
        folders = push_recent_folder(folders, str(folder))

    assert folders == list(reversed(created[1:]))


def test異常系_push_recent_folderメソッドが存在しないパスを追加しないこと(tmp_path):
    assert push_recent_folder([], str(tmp_path / "missing")) == []


def test正常系_existing_recent_foldersメソッドが無いフォルダーを除くこと(tmp_path):
    kept = tmp_path / "kept"
    kept.mkdir()

    assert existing_recent_folders([str(kept), str(tmp_path / "gone")]) == [str(kept)]
