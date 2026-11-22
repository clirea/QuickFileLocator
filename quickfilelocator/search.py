"""フォルダーを走査して、検索条件に一致するファイルを探す。"""

from __future__ import annotations

import os
from collections.abc import Callable, Mapping
from dataclasses import dataclass
from pathlib import Path

from quickfilelocator.query import SearchQuery
from quickfilelocator.readers import FileReader, UnreadableFileError

# これを超えるファイルは読み込まずに除外する。全文をメモリに載せるための上限。
DEFAULT_MAX_FILE_BYTES = 100 * 1024 * 1024


class SearchDirectoryError(Exception):
    """検索先フォルダーを使えないときに送出する。メッセージはそのまま画面に表示する。"""


@dataclass(frozen=True)
class ExcludedFile:
    """検索できなかったファイル(またはフォルダー)と、その理由。"""

    path: Path
    reason: str


@dataclass(frozen=True)
class SearchSummary:
    """1回の検索の結果。"""

    target_count: int  # 検索対象の拡張子を持つファイル数
    searched_count: int  # 本文を読んで照合できたファイル数
    matched_count: int
    excluded: tuple[ExcludedFile, ...]
    cancelled: bool


class FileSearcher:
    """拡張子ごとの読み取り手段を使って、フォルダー配下を全文検索する。"""

    def __init__(
        self,
        readers: Mapping[str, FileReader],
        max_file_bytes: int = DEFAULT_MAX_FILE_BYTES,
    ) -> None:
        self._readers = {suffix.lower(): reader for suffix, reader in readers.items()}
        self._max_file_bytes = max_file_bytes

    def search(
        self,
        query: SearchQuery,
        directory: Path,
        *,
        on_match: Callable[[Path], None],
        on_progress: Callable[[int, int], None],
        is_cancelled: Callable[[], bool] = lambda: False,
    ) -> SearchSummary:
        """directory 配下を検索する。

        一致するたびに on_match(パス)、1ファイル処理するたびに
        on_progress(処理済み件数, 対象件数) を呼ぶ。is_cancelled が真を返すと、
        そこまでの結果を cancelled=True で返す。
        """
        if not directory.is_dir():
            raise SearchDirectoryError(f"検索先フォルダーが見つかりません: {directory}")

        excluded: list[ExcludedFile] = []
        targets = self._collect_targets(directory, excluded, is_cancelled)
        searched_count = 0
        matched_count = 0
        cancelled = is_cancelled()

        if not cancelled:
            on_progress(0, len(targets))
            for processed_count, path in enumerate(targets, start=1):
                if is_cancelled():
                    cancelled = True
                    break
                try:
                    content = self._read(path)
                except (UnreadableFileError, OSError) as error:
                    excluded.append(ExcludedFile(path, self._describe(error)))
                else:
                    searched_count += 1
                    if query.matches(content):
                        matched_count += 1
                        on_match(path)
                on_progress(processed_count, len(targets))

        return SearchSummary(
            target_count=len(targets),
            searched_count=searched_count,
            matched_count=matched_count,
            excluded=tuple(excluded),
            cancelled=cancelled,
        )

    def _collect_targets(
        self,
        directory: Path,
        excluded: list[ExcludedFile],
        is_cancelled: Callable[[], bool],
    ) -> list[Path]:
        """対象拡張子のファイルを、結果が毎回同じ順になるよう名前順で集める。"""

        def on_walk_error(error: OSError) -> None:
            excluded.append(ExcludedFile(Path(error.filename or directory), self._describe(error)))

        targets: list[Path] = []
        for root, dir_names, file_names in os.walk(directory, onerror=on_walk_error):
            if is_cancelled():
                break
            dir_names.sort()
            targets.extend(
                Path(root, name) for name in sorted(file_names) if os.path.splitext(name)[1].lower() in self._readers
            )
        return targets

    def _read(self, path: Path) -> str:
        size = path.stat().st_size
        if size > self._max_file_bytes:
            limit_mb = self._max_file_bytes / (1024 * 1024)
            raise UnreadableFileError(f"サイズが上限({limit_mb:g}MB)を超えています")
        return self._readers[path.suffix.lower()].read(path)

    @staticmethod
    def _describe(error: Exception) -> str:
        if isinstance(error, OSError):
            return f"読み取れません ({error.strerror or type(error).__name__})"
        return str(error)
