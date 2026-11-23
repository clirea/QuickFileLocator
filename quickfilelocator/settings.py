"""画面状態の保存に使うキーと、最近使ったフォルダーの純関数。"""

from __future__ import annotations

from pathlib import Path

ORGANIZATION_NAME = "QuickFileLocator"
APPLICATION_NAME = "QuickFileLocator"

KEY_GEOMETRY = "geometry"
KEY_SPLITTER = "splitter"
KEY_DIRECTORY = "directory"
KEY_RECENT_FOLDERS = "recentFolders"
KEY_CASE_SENSITIVE = "caseSensitive"
KEY_REGEX = "regex"
KEY_SORT_MODE = "sortMode"

MAX_RECENT_FOLDERS = 10
SORT_BY_NAME = "名前"
SORT_BY_HITS = "ヒット件数"
SORT_BY_PATH = "パス"
SORT_MODES = (SORT_BY_NAME, SORT_BY_HITS, SORT_BY_PATH)
DEFAULT_SORT_MODE = SORT_BY_NAME


def as_bool(value: object, default: bool) -> bool:
    """QSettings が返しうる真偽値表現を bool にする。"""
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return bool(value)
    text = str(value).strip().lower()
    if text in {"true", "1", "yes"}:
        return True
    if text in {"false", "0", "no", ""}:
        return False
    return default


def as_string_list(value: object) -> list[str]:
    """QSettings が返しうる文字列リストをそろえる。"""
    if value is None:
        return []
    if isinstance(value, str):
        return [value] if value else []
    if isinstance(value, (list, tuple)):
        return [str(item) for item in value if str(item)]
    return [str(value)]


def existing_directory(value: object) -> Path | None:
    """存在するフォルダーだけを Path として返す。"""
    if value is None or str(value).strip() == "":
        return None
    path = Path(str(value)).expanduser()
    if path.is_dir():
        return path
    return None


def existing_recent_folders(value: object) -> list[str]:
    """最近使ったフォルダーのうち、今も存在するものを新しい順で最大件数まで返す。"""
    folders: list[str] = []
    seen: set[str] = set()
    for item in as_string_list(value):
        path = existing_directory(item)
        if path is None:
            continue
        text = str(path)
        if text in seen:
            continue
        folders.append(text)
        seen.add(text)
        if len(folders) >= MAX_RECENT_FOLDERS:
            break
    return folders


def push_recent_folder(folders: list[str], directory: str) -> list[str]:
    """存在するフォルダーを先頭へ移し、最大件数に収める。"""
    path = existing_directory(directory)
    if path is None:
        return list(folders)
    text = str(path)
    updated = [item for item in folders if item != text]
    updated.insert(0, text)
    return updated[:MAX_RECENT_FOLDERS]
