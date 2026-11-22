import pytest

from quickfilelocator.query import QuerySyntaxError, SearchQuery


@pytest.mark.parametrize(
    ("text", "expected_groups"),
    [
        ("foo", (("foo",),)),
        ("foo && bar", (("foo",), ("bar",))),
        ("foo&&bar", (("foo",), ("bar",))),
        ("foo || bar", (("foo", "bar"),)),
        ("foo bar", (("foo", "bar"),)),
        ("foo\u3000bar", (("foo", "bar"),)),
        ("foo || bar && baz", (("foo", "bar"), ("baz",))),
        ('"foo bar"', (("foo bar",),)),
        ('"foo bar" && baz', (("foo bar",), ("baz",))),
        ('"a && b || c"', (("a && b || c",),)),
        ('pre"fix word"', (("prefix word",),)),
        ('\\"foo\\"', (('"foo"',),)),
        ("C:\\Users", (("C:\\Users",),)),
        ("foo &&", (("foo",),)),
        ("&& foo || || bar", (("foo", "bar"),)),
        ('foo ""', (("foo",),)),
    ],
)
def test正常系_parseメソッドが検索式をORグループのANDへ分解すること(text, expected_groups):
    assert SearchQuery.parse(text).groups == expected_groups


def test正常系_parseメソッドが既定で大文字小文字を区別する条件を返すこと():
    assert SearchQuery.parse("Foo").case_sensitive is True


def test正常系_parseメソッドが区別しない指定のとき検索語を小文字にそろえること():
    assert SearchQuery.parse("Foo ÄB", case_sensitive=False).groups == (("foo", "äb"),)


@pytest.mark.parametrize("text", ["", "   ", "&&", "||", "&& || &&", '""'])
def test異常系_parseメソッドが検索語のない式でQuerySyntaxErrorを送出すること(text):
    with pytest.raises(QuerySyntaxError, match="検索する語"):
        SearchQuery.parse(text)


@pytest.mark.parametrize("text", ['"foo', 'foo "bar baz', 'a && "b'])
def test異常系_parseメソッドが閉じていない引用符でQuerySyntaxErrorを送出すること(text):
    with pytest.raises(QuerySyntaxError, match="引用符"):
        SearchQuery.parse(text)


@pytest.mark.parametrize(
    ("text", "content", "expected"),
    [
        ("foo", "xx foo yy", True),
        ("foo", "bar", False),
        ("foo && bar", "bar and foo", True),
        ("foo && bar", "foo only", False),
        ("foo || bar", "bar only", True),
        ("foo || bar", "baz", False),
        ("foo bar", "bar only", True),
        ("foo || bar && baz", "bar baz", True),
        ("foo || bar && baz", "foo bar", False),
        ('"foo bar"', "a foo bar b", True),
        ('"foo bar"', "foo  bar", False),
        ('"foo bar"', "bar foo", False),
        ('"foo"', "foo", True),
        ('\\"foo\\"', "foo", False),
        ('\\"foo\\"', 'say "foo"', True),
        ("検索 && 文書", "文書を検索する", True),
        ("Foo", "foo", False),
    ],
)
def test正常系_matchesメソッドが検索条件どおりに本文を判定すること(text, content, expected):
    assert SearchQuery.parse(text).matches(content) is expected


@pytest.mark.parametrize(
    ("text", "content", "expected"),
    [
        ("Foo", "xx FOO yy", True),
        ("foo && BAR", "Bar Foo", True),
        ('"Foo Bar"', "foo bar", True),
        ("foo", "bar", False),
    ],
)
def test正常系_matchesメソッドが区別しない指定のとき大文字小文字を無視して判定すること(text, content, expected):
    assert SearchQuery.parse(text, case_sensitive=False).matches(content) is expected


def test正常系_find_matchesメソッドが複数箇所の件数を返すこと():
    assert len(SearchQuery.parse("ab").find_matches("ab x ab")) == 2


def test正常系_find_matchesメソッドが1件目の開始位置を返すこと():
    assert SearchQuery.parse("ab").find_matches("ab x ab")[0] == (0, 2)


def test正常系_find_matchesメソッドが2件目の開始位置を返すこと():
    assert SearchQuery.parse("ab").find_matches("ab x ab")[1] == (5, 7)


def test正常系_find_matchesメソッドが条件を満たさない本文では空を返すこと():
    assert SearchQuery.parse("foo && bar").find_matches("foo only") == ()


def test正常系_unique_termsメソッドが重複する語を1つにすること():
    assert SearchQuery.parse("foo || foo && bar").unique_terms() == ("foo", "bar")


def test正常系_matchesメソッドが正規表現の繰り返しに一致すること():
    assert SearchQuery.parse("fo+", regex=True).matches("foo") is True


def test正常系_matchesメソッドが正規表現に合わない本文を拒否すること():
    assert SearchQuery.parse("fo+", regex=True).matches("fx") is False


def test正常系_matchesメソッドが正規表現でもANDの両方を要求すること():
    assert SearchQuery.parse("fo+ && ba+", regex=True).matches("foo") is False


def test正常系_matchesメソッドが引用符付きの正規表現文字を文字列として扱うこと():
    assert SearchQuery.parse('"fo+"', regex=True).matches("foo") is False


def test異常系_parseメソッドが不正な正規表現でQuerySyntaxErrorを送出すること():
    with pytest.raises(QuerySyntaxError, match="正規表現が不正です"):
        SearchQuery.parse("(", regex=True)


def test正常系_matchesメソッドが正規表現で大文字小文字を無視すること():
    assert SearchQuery.parse("fo+", case_sensitive=False, regex=True).matches("FOO") is True


def test正常系_find_matchesメソッドが正規表現の一致範囲を返すこと():
    assert SearchQuery.parse("fo+", regex=True).find_matches("x foo y") == ((2, 5),)


def test正常系_find_matchesメソッドが正規表現オフのときメタ文字を文字列として扱うこと():
    assert SearchQuery.parse("fo+", regex=False).find_matches("foo") == ()


def test正常系_parseメソッドが正規表現指定を保持すること():
    assert SearchQuery.parse("foo", regex=True).regex is True


def test正常系_parseメソッドが既定で正規表現を使わないこと():
    assert SearchQuery.parse("foo").regex is False
