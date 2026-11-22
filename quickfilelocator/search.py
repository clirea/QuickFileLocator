"""フォルダーを走査して、検索条件に一致するファイルを探す。"""

from __future__ import annotations

import os
import threading
from collections.abc import Callable, Mapping
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from pathlib import Path

from quickfilelocator.query import SearchQuery
from quickfilelocator.readers import FileReader, UnreadableFileError

# これを超えるファイルは読み込まずに除外する。全文をメモリに載せるための上限。
DEFAULT_MAX_FILE_BYTES = 100 * 1024 * 1024


def first_hit_snippet(content: str, hits: tuple[tuple[int, int], ...]) -> str:
    """最初のヒットを含む行を、折り返さず1行の抜粋として返す。"""
    start, _end = hits[0]
    line_start = content.rfind("\n", 0, start) + 1
    line_end = content.find("\n", start)
    if line_end < 0:
        line_end = len(content)
    return content[line_start:line_end].replace("\r", "")


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
        max_workers: int | None = None,
    ) -> None:
        self._readers = {suffix.lower(): reader for suffix, reader in readers.items()}
        self._max_file_bytes = max_file_bytes
        self._max_workers = min(8, os.cpu_count() or 1) if max_workers is None else max_workers

    def search(
        self,
        query: SearchQuery,
        directory: Path,
        *,
        on_match: Callable[[Path, int, str], None],
        on_progress: Callable[[int, int], None],
        is_cancelled: Callable[[], bool] = lambda: False,
    ) -> SearchSummary:
        """directory 配下を検索する。

        一致するたびに on_match(パス, ヒット件数, 抜粋)、1ファイル処理するたびに
        on_progress(処理済み件数, 対象件数) を呼ぶ。本文ヒットがあれば抜粋は最初の
        ヒットを含む行を1行にしたもの、ファイル名だけ一致した場合はファイル名である。
        is_cancelled が真を返すと、そこまでの結果を cancelled=True で返す。
        """
        if not directory.is_dir():
            raise SearchDirectoryError(f"検索先フォルダーが見つかりません: {directory}")

        excluded: list[ExcludedFile] = []
        targets = self._collect_targets(directory, excluded, is_cancelled)
        searched_count = 0
        matched_count = 0
        processed_count = 0
        cancelled = is_cancelled()

        if not cancelled:
            target_count = len(targets)
            on_progress(0, target_count)
            lock = threading.Lock()

            def process_file(path: Path) -> None:
                nonlocal searched_count, matched_count, processed_count, cancelled
                if is_cancelled():
                    with lock:
                        cancelled = True
                    return
                try:
                    content = self.read(path)
                except (UnreadableFileError, OSError) as error:
                    reason = self._describe(error)
                    with lock:
                        excluded.append(ExcludedFile(path, reason))
                        processed_count += 1
                        on_progress(processed_count, target_count)
                    return

                hits = query.find_matches(content)
                match_args: tuple[Path, int, str] | None
                if hits:
                    match_args = (path, len(hits), first_hit_snippet(content, hits))
                elif query.matches(path.name):
                    match_args = (path, 1, path.name)
                else:
                    match_args = None

                with lock:
                    searched_count += 1
                    if match_args is not None:
                        matched_count += 1
                        on_match(*match_args)
                    processed_count += 1
                    on_progress(processed_count, target_count)

            with ThreadPoolExecutor(max_workers=self._max_workers) as executor:
                futures = []
                for path in targets:
                    if is_cancelled():
                        cancelled = True
                        break
                    futures.append(executor.submit(process_file, path))
                for future in as_completed(futures):
                    if is_cancelled():
                        cancelled = True
                    future.result()
                if is_cancelled():
                    cancelled = True

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

    def read(self, path: Path) -> str:
        """1ファイルの本文を返す。プレビュー表示でも使う。"""
        size = path.stat().st_size
        if size > self._max_file_bytes:
            limit_mb = self._max_file_bytes / (1024 * 1024)
            raise UnreadableFileError(f"サイズが上限({limit_mb:g}MB)を超えています")
        suffix = path.suffix.lower()
        if suffix not in self._readers:
            raise UnreadableFileError(f"未対応の拡張子です ({suffix or 'なし'})")
        return self._readers[suffix].read(path)

    @staticmethod
    def _describe(error: Exception) -> str:
        if isinstance(error, OSError):
            return f"読み取れません ({error.strerror or type(error).__name__})"
        return str(error)
