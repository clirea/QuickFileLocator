from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

from quickfilelocator import app
from quickfilelocator.app import default_search_directory, main


@pytest.fixture
def cwd(tmp_path):
    return tmp_path / "cwd"


def test正常系_default_search_directoryメソッドが引数で指定されたフォルダーを返すこと(cwd):
    result = default_search_directory(["/data/docs"], frozen=True, executable="/opt/app/app.exe", cwd=cwd)

    assert result == Path("/data/docs")


def test正常系_default_search_directoryメソッドがexe化時にexeのあるフォルダーを返すこと(tmp_path, cwd):
    executable = tmp_path / "dist" / "QuickFileLocator.exe"

    result = default_search_directory([], frozen=True, executable=str(executable), cwd=cwd)

    assert result == executable.resolve().parent


def test正常系_default_search_directoryメソッドがスクリプト実行時にカレントディレクトリを返すこと(cwd):
    result = default_search_directory([], frozen=False, executable="/usr/bin/python3", cwd=cwd)

    assert result == cwd


@pytest.fixture
def launched():
    """QApplication と MainWindow を差し替えて main を実行できるようにする。"""
    with (
        patch.object(app, "QApplication") as application_class,
        patch.object(app, "MainWindow") as window_class,
    ):
        application_class.return_value.arguments.return_value = ["main.py", "/data/docs"]
        application_class.return_value.exec_.return_value = 7
        yield MagicMock(application_class=application_class, window_class=window_class)


def test正常系_mainメソッドが起動引数のフォルダーを初期の検索先として画面へ渡すこと(launched):
    main(["main.py", "/data/docs"])

    assert launched.window_class.call_args.args[1] == Path("/data/docs")


def test正常系_mainメソッドが画面を表示すること(launched):
    main(["main.py"])

    launched.window_class.return_value.show.assert_called_once_with()


def test正常系_mainメソッドがイベントループの終了コードを返すこと(launched):
    assert main(["main.py"]) == 7
