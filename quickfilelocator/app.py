"""アプリケーションの組み立てと起動。"""

from __future__ import annotations

import sys
from collections.abc import Sequence
from pathlib import Path

from PyQt5.QtWidgets import QApplication

from quickfilelocator.gui import MainWindow
from quickfilelocator.readers import default_readers
from quickfilelocator.search import FileSearcher


def default_search_directory(arguments: Sequence[str], *, frozen: bool, executable: str, cwd: Path) -> Path:
    """起動時に検索先欄へ入れるフォルダーを決める。

    引数で指定されていればそれを使う。exe化して配布した場合は、従来どおり
    exeを置いたフォルダーを検索先とする。
    """
    if arguments:
        return Path(arguments[0]).expanduser()
    if frozen:
        return Path(executable).resolve().parent
    return cwd


def main(argv: Sequence[str] | None = None) -> int:
    app = QApplication(list(sys.argv if argv is None else argv))
    directory = default_search_directory(
        app.arguments()[1:],
        frozen=bool(getattr(sys, "frozen", False)),
        executable=sys.executable,
        cwd=Path.cwd(),
    )
    main_window = MainWindow(FileSearcher(default_readers()), directory)
    main_window.show()
    return app.exec_()
