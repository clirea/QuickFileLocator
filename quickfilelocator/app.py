"""アプリケーションの組み立てと起動。"""

from __future__ import annotations

import sys
from collections.abc import Sequence
from pathlib import Path

from PyQt6.QtCore import QSettings
from PyQt6.QtWidgets import QApplication

from quickfilelocator.gui import MainWindow
from quickfilelocator.readers import default_readers
from quickfilelocator.search import FileSearcher
from quickfilelocator.settings import APPLICATION_NAME, KEY_DIRECTORY, ORGANIZATION_NAME, existing_directory


def default_search_directory(
    arguments: Sequence[str],
    *,
    frozen: bool,
    executable: str,
    cwd: Path,
    last_directory: Path | str | None = None,
) -> Path:
    """起動時に検索先欄へ入れるフォルダーを決める。

    引数で指定されていればそれを使う。無ければ前回の検索先（存在するとき）。
    exe化して配布した場合は、従来どおり exeを置いたフォルダーを検索先とする。
    """
    if arguments:
        return Path(arguments[0]).expanduser()
    saved = existing_directory(last_directory)
    if saved is not None:
        return saved
    if frozen:
        return Path(executable).resolve().parent
    return cwd


def main(argv: Sequence[str] | None = None) -> int:
    app = QApplication(list(sys.argv if argv is None else argv))
    app.setOrganizationName(ORGANIZATION_NAME)
    app.setApplicationName(APPLICATION_NAME)
    settings = QSettings()
    directory = default_search_directory(
        app.arguments()[1:],
        frozen=bool(getattr(sys, "frozen", False)),
        executable=sys.executable,
        cwd=Path.cwd(),
        last_directory=settings.value(KEY_DIRECTORY),
    )
    main_window = MainWindow(FileSearcher(default_readers()), directory, settings=settings)
    main_window.show()
    return app.exec()
