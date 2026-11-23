"""検索対象ファイルから本文テキストを取り出す。"""

from __future__ import annotations

import codecs
from collections.abc import Callable, Iterator
from pathlib import Path
from typing import Any, Protocol

import chardet
import docx
import pypdf

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
_WML_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_TXBX_CONTENT_TAG = f"{{{_WML_NS}}}txbxContent"
_PARAGRAPH_TAG = f"{{{_WML_NS}}}p"
_TEXT_TAG = f"{{{_WML_NS}}}t"
_VML_TEXTBOX_TAG = "{urn:schemas-microsoft-com:vml}textbox"
_HEADER_FOOTER_ATTRS = (
    "header",
    "footer",
    "first_page_header",
    "first_page_footer",
    "even_page_header",
    "even_page_footer",
)


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
    """Word文書(.docx)の本文・表・ヘッダー・フッター・テキストボックスのテキストを読む。"""

    def __init__(self, open_document: Callable[[str], Any] = docx.Document) -> None:
        self._open_document = open_document

    def read(self, path: Path) -> str:
        if path.name.startswith(WORD_LOCK_FILE_PREFIX):
            raise UnreadableFileError("Wordが編集中に作成する一時ファイルです")
        try:
            document = self._open_document(str(path))
            return "\n".join(self._iter_document_texts(document))
        except OSError:
            raise
        except Exception as error:
            # python-docx は壊れた文書に対して zipfile / lxml などの多様な例外をそのまま返す
            raise UnreadableFileError(f"Word文書として開けません ({type(error).__name__})") from error

    def _iter_document_texts(self, document: Any) -> Iterator[str]:
        yield from self._iter_texts(document)
        for text in self._iter_header_footer_texts(document):
            if text.strip():
                yield text
        yield from self._iter_textbox_texts(document)

    def _iter_texts(self, container: Any) -> Iterator[str]:
        """段落と表を出現順にたどる。表のセルは入れ子の表も含めて再帰する。"""
        for block in container.iter_inner_content():
            if hasattr(block, "rows"):
                for row in block.rows:
                    for cell in row.cells:
                        yield from self._iter_texts(cell)
            else:
                yield block.text

    def _iter_header_footer_texts(self, document: Any) -> Iterator[str]:
        for header_footer in self._iter_header_footer_containers(document):
            yield from self._iter_texts(header_footer)

    def _iter_header_footer_containers(self, document: Any) -> Iterator[Any]:
        for section in self._iter_sections(document):
            for attr in _HEADER_FOOTER_ATTRS:
                header_footer = getattr(section, attr, None)
                if header_footer is None:
                    continue
                if getattr(header_footer, "is_linked_to_previous", False):
                    continue
                yield header_footer

    def _iter_sections(self, document: Any) -> Iterator[Any]:
        sections = getattr(document, "sections", None)
        if sections is None:
            return
        try:
            yield from sections
        except TypeError:
            return

    def _iter_textbox_texts(self, document: Any) -> Iterator[str]:
        for root in self._iter_xml_roots(document):
            yield from self._iter_txbx_content_texts(root)
            yield from self._iter_vml_textbox_texts(root)

    def _iter_xml_roots(self, document: Any) -> Iterator[Any]:
        seen: set[int] = set()
        for candidate in (getattr(document, "element", None),):
            yield from self._yield_xml_root(candidate, seen)
        for header_footer in self._iter_header_footer_containers(document):
            element = getattr(header_footer, "element", None)
            if element is None:
                element = getattr(header_footer, "_element", None)
            yield from self._yield_xml_root(element, seen)

    def _yield_xml_root(self, element: Any, seen: set[int]) -> Iterator[Any]:
        if not isinstance(getattr(element, "tag", None), str):
            return
        identity = id(element)
        if identity in seen:
            return
        seen.add(identity)
        yield element

    def _iter_txbx_content_texts(self, root: Any) -> Iterator[str]:
        for box in self._iter_tagged(root, _TXBX_CONTENT_TAG):
            for paragraph in self._iter_paragraphs_skipping_nested_txbx(box):
                text = self._xml_text_nodes(paragraph)
                if text.strip():
                    yield text

    def _iter_vml_textbox_texts(self, root: Any) -> Iterator[str]:
        for box in self._iter_tagged(root, _VML_TEXTBOX_TAG):
            if any(self._iter_tagged(box, _TXBX_CONTENT_TAG)):
                continue
            text = self._xml_text_nodes(box)
            if text.strip():
                yield text

    def _iter_paragraphs_skipping_nested_txbx(self, node: Any) -> Iterator[Any]:
        for child in self._xml_children(node):
            tag = getattr(child, "tag", None)
            if tag == _TXBX_CONTENT_TAG:
                continue
            if tag == _PARAGRAPH_TAG:
                yield child
            else:
                yield from self._iter_paragraphs_skipping_nested_txbx(child)

    def _xml_text_nodes(self, node: Any) -> str:
        parts: list[str] = []
        for child in self._iter_tagged(node, _TEXT_TAG):
            text = getattr(child, "text", None)
            if isinstance(text, str):
                parts.append(text)
        return "".join(parts)

    def _iter_tagged(self, node: Any, tag: str) -> Iterator[Any]:
        if not isinstance(getattr(node, "tag", None), str):
            return
        iter_elements = getattr(node, "iter", None)
        if not callable(iter_elements):
            return
        try:
            found = iter_elements(tag)
        except TypeError:
            return
        for element in found:
            if isinstance(getattr(element, "tag", None), str):
                yield element
            else:
                return

    def _xml_children(self, node: Any) -> Iterator[Any]:
        try:
            children = list(node)
        except TypeError:
            return
        for child in children:
            if isinstance(getattr(child, "tag", None), str):
                yield child


class PdfFileReader:
    """PDFからテキストを取り出す。抽出できないファイルは除外対象にする。"""

    def __init__(self, open_document: Callable[[str], Any] = pypdf.PdfReader) -> None:
        self._open_document = open_document

    def read(self, path: Path) -> str:
        try:
            reader = self._open_document(str(path))
            if getattr(reader, "is_encrypted", False) is True:
                raise UnreadableFileError("PDFからテキストを取り出せません (暗号化されています)")
            texts = [self._page_text(page) for page in self._pages(reader)]
            joined = "\n".join(texts)
            if not joined.strip():
                raise UnreadableFileError("PDFからテキストを取り出せません (テキストを抽出できません)")
            return joined
        except UnreadableFileError:
            raise
        except OSError:
            raise
        except Exception as error:
            raise UnreadableFileError(f"PDFからテキストを取り出せません ({type(error).__name__})") from error

    def _pages(self, reader: Any) -> Iterator[Any]:
        pages = getattr(reader, "pages", None)
        if pages is None:
            return
        try:
            yield from pages
        except TypeError:
            return

    def _page_text(self, page: Any) -> str:
        extract = getattr(page, "extract_text", None)
        if not callable(extract):
            return ""
        text = extract()
        return text if isinstance(text, str) else ""


def default_readers() -> dict[str, FileReader]:
    """拡張子(小文字)ごとの既定の読み取り手段を返す。"""
    text_reader = TextFileReader()
    return {
        ".docx": DocxFileReader(),
        ".txt": text_reader,
        ".md": text_reader,
        ".markdown": text_reader,
        ".csv": text_reader,
        ".pdf": PdfFileReader(),
    }
