"""検索対象ファイルから本文テキストを取り出す。"""

from __future__ import annotations

import codecs
from collections.abc import Callable, Iterator
from pathlib import Path
from typing import Any, Protocol

import chardet
import docx

# 文字コード判定に渡す先頭バイト数。巨大なファイルでも判定時間を一定に抑える。
DETECTION_SAMPLE_BYTES = 256 * 1024
# 日本語Windowsの既定の文字コード。UTF-8でないテキストはまずこれとして扱う。
FALLBACK_ENCODINGS = ("cp932",)
# 判定結果がこれらのときだけ、FALLBACK_ENCODINGS より先に試す。短いシフトJISの文章は
# 欧文やヘブライ語などの1バイト文字コードと誤判定されやすく、その結果でも復号自体は
# 成功して文字化けした本文になるため、判定結果を一律には信用しない。
TRUSTED_DETECTIONS = frozenset(
    {"euc_jp", "iso2022_jp", "utf-16", "utf-16-le", "utf-16-be", "utf-32", "utf-32-le", "utf-32-be"}
)
# 先頭にあれば文字コードが確定するBOM。UTF-32はUTF-16と前方一致するため先に調べる。
BOM_ENCODINGS = (
    (codecs.BOM_UTF8, "utf-8-sig"),
    (codecs.BOM_UTF32_LE, "utf-32"),
    (codecs.BOM_UTF32_BE, "utf-32"),
    (codecs.BOM_UTF16_LE, "utf-16"),
    (codecs.BOM_UTF16_BE, "utf-16"),
)
# ISO-2022-JP(JIS)は7ビットのためUTF-8としても復号できてしまう。漢字への切り替えを示す
# このエスケープシーケンスがあれば、UTF-8より先に試す。
ISO2022_JP_KANJI_ESCAPE = b"\x1b$"
WORD_LOCK_FILE_PREFIX = "~$"


class UnreadableFileError(Exception):
    """本文を取り出せないファイルに対して送出する。メッセージは除外理由として表示する。"""


class FileReader(Protocol):
    def read(self, path: Path) -> str:
        """本文を返す。取り出せない場合は UnreadableFileError を送出する。"""
        ...


class TextFileReader:
    """文字コードを判定してテキストファイルを読む。"""

    def __init__(self, detect: Callable[[bytes], dict[str, Any]] = chardet.detect) -> None:
        self._detect = detect

    def read(self, path: Path) -> str:
        return self.decode(path.read_bytes())

    def decode(self, data: bytes) -> str:
        """BOM、JIS、UTF-8、信用できる判定結果、CP932、その他の判定結果の順に復号を試みる。"""
        for encoding in self._candidate_encodings(data):
            try:
                return data.decode(encoding)
            except UnicodeDecodeError:
                continue
        raise UnreadableFileError("文字コードを判定できません")

    def _candidate_encodings(self, data: bytes) -> Iterator[str]:
        for bom, encoding in BOM_ENCODINGS:
            if data.startswith(bom):
                yield encoding
                return
        if ISO2022_JP_KANJI_ESCAPE in data:
            yield "iso2022_jp"
        yield "utf-8"
        detected = self._detect_encoding(data)
        if detected in TRUSTED_DETECTIONS:
            yield detected
        yield from FALLBACK_ENCODINGS
        if detected:
            yield detected

    def _detect_encoding(self, data: bytes) -> str | None:
        """判定された文字コードを、Pythonの正規化した名前で返す。"""
        detected = self._detect(data[:DETECTION_SAMPLE_BYTES]).get("encoding")
        try:
            return codecs.lookup(detected).name if detected else None
        except LookupError:
            return None


class DocxFileReader:
    """Word文書(.docx)の本文と表のテキストを読む。"""

    def __init__(self, open_document: Callable[[str], Any] = docx.Document) -> None:
        self._open_document = open_document

    def read(self, path: Path) -> str:
        if path.name.startswith(WORD_LOCK_FILE_PREFIX):
            raise UnreadableFileError("Wordが編集中に作成する一時ファイルです")
        try:
            document = self._open_document(str(path))
            return "\n".join(self._iter_texts(document))
        except OSError:
            raise
        except Exception as error:
            # python-docx は壊れた文書に対して zipfile / lxml などの多様な例外をそのまま返す
            raise UnreadableFileError(f"Word文書として開けません ({type(error).__name__})") from error

    def _iter_texts(self, container: Any) -> Iterator[str]:
        """段落と表を出現順にたどる。表のセルは入れ子の表も含めて再帰する。"""
        for block in container.iter_inner_content():
            if hasattr(block, "rows"):
                for row in block.rows:
                    for cell in row.cells:
                        yield from self._iter_texts(cell)
            else:
                yield block.text


def default_readers() -> dict[str, FileReader]:
    """拡張子(小文字)ごとの既定の読み取り手段を返す。"""
    return {".docx": DocxFileReader(), ".txt": TextFileReader()}
