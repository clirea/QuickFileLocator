"""検索式の解析と、本文との照合。"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

AND_OPERATOR = "&&"
OR_OPERATOR = "||"
QUOTE = '"'
ESCAPED_QUOTE = '\\"'


class QuerySyntaxError(ValueError):
    """検索式を解釈できないときに送出する。メッセージはそのまま画面に表示する。"""


def _unique_terms(groups: tuple[tuple[str, ...], ...]) -> tuple[str, ...]:
    seen: list[str] = []
    for group in groups:
        for term in group:
            if term and term not in seen:
                seen.append(term)
    return tuple(seen)


def _compile_patterns(terms: tuple[str, ...], *, case_sensitive: bool) -> tuple[re.Pattern[str], ...]:
    flags = 0 if case_sensitive else re.IGNORECASE
    compiled: list[re.Pattern[str]] = []
    for term in terms:
        try:
            compiled.append(re.compile(term, flags))
        except re.error as error:
            raise QuerySyntaxError(f"正規表現が不正です: {term}: {error}") from error
    return tuple(compiled)


@dataclass(frozen=True)
class SearchQuery:
    """ORグループをANDで結んだ検索条件。

    groups の各要素が1つのORグループで、すべてのグループが成立した本文だけが一致する。
    """

    groups: tuple[tuple[str, ...], ...]
    case_sensitive: bool = True
    regex: bool = False
    compiled_patterns: tuple[re.Pattern[str], ...] = field(default=(), repr=False, compare=False)

    @classmethod
    def parse(cls, text: str, *, case_sensitive: bool = True, regex: bool = False) -> SearchQuery:
        """検索式を解析する。

        - ``&&`` はAND、``||`` と空白はOR
        - ``"..."`` で囲んだ範囲は、空白や演算子を含めて1つの語(フレーズ)として扱う
        - 引用符そのものを検索するときは ``\\"`` と書く
        - ``regex=True`` のとき、引用符なしの語は正規表現、引用符内は文字列として扱う
        """
        groups: list[tuple[str, ...]] = []
        terms: list[str] = []
        buffer: list[str] = []
        quoted = False

        def flush_term() -> None:
            term = "".join(buffer)
            buffer.clear()
            if term:
                if not regex and not case_sensitive:
                    term = term.casefold()
                terms.append(term)

        def flush_group() -> None:
            flush_term()
            if terms:
                groups.append(tuple(terms))
            terms.clear()

        def append_char(char: str) -> None:
            buffer.append(re.escape(char) if regex and quoted else char)

        index = 0
        while index < len(text):
            if text.startswith(ESCAPED_QUOTE, index):
                append_char(QUOTE)
                index += len(ESCAPED_QUOTE)
            elif text[index] == QUOTE:
                quoted = not quoted
                index += 1
            elif not quoted and text.startswith(AND_OPERATOR, index):
                flush_group()
                index += len(AND_OPERATOR)
            elif not quoted and text.startswith(OR_OPERATOR, index):
                flush_term()
                index += len(OR_OPERATOR)
            elif not quoted and text[index].isspace():
                flush_term()
                index += 1
            else:
                append_char(text[index])
                index += 1

        if quoted:
            raise QuerySyntaxError('引用符(")が閉じられていません。')
        flush_group()
        if not groups:
            raise QuerySyntaxError("検索する語を入力してください。")
        groups_tuple = tuple(groups)
        compiled = _compile_patterns(_unique_terms(groups_tuple), case_sensitive=case_sensitive) if regex else ()
        return cls(
            groups=groups_tuple,
            case_sensitive=case_sensitive,
            regex=regex,
            compiled_patterns=compiled,
        )

    def matches(self, content: str) -> bool:
        """本文がこの検索条件を満たすかを返す。"""
        if not self.regex:
            haystack = content if self.case_sensitive else content.casefold()
            return all(any(term in haystack for term in group) for group in self.groups)
        compiled = self._compiled_by_term()
        return all(any(compiled[term].search(content) is not None for term in group) for group in self.groups)

    def unique_terms(self) -> tuple[str, ...]:
        """検索式に現れる語を、重複なしで出現順に返す。"""
        return _unique_terms(self.groups)

    def find_matches(self, content: str) -> tuple[tuple[int, int], ...]:
        """条件を満たす本文について、検索語の出現範囲 (start, end) を開始位置順で返す。"""
        if not self.matches(content):
            return ()
        if not self.regex:
            flags = 0 if self.case_sensitive else re.IGNORECASE
            spans: list[tuple[int, int]] = []
            seen: set[tuple[int, int]] = set()
            for term in self.unique_terms():
                for match in re.finditer(re.escape(term), content, flags):
                    span = (match.start(), match.end())
                    if span not in seen:
                        seen.add(span)
                        spans.append(span)
            return tuple(sorted(spans))
        spans = []
        seen: set[tuple[int, int]] = set()
        compiled = self._compiled_by_term()
        for term in self.unique_terms():
            for match in compiled[term].finditer(content):
                start, end = match.start(), match.end()
                if start == end:
                    continue
                span = (start, end)
                if span not in seen:
                    seen.add(span)
                    spans.append(span)
        return tuple(sorted(spans))

    def _compiled_by_term(self) -> dict[str, re.Pattern[str]]:
        terms = self.unique_terms()
        patterns = self.compiled_patterns
        if len(patterns) != len(terms):
            patterns = _compile_patterns(terms, case_sensitive=self.case_sensitive)
        return dict(zip(terms, patterns, strict=True))
