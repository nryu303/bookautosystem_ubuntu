"""code_normalizer のユニットテスト"""

import pytest
from src.parser.code_normalizer import (
    normalize_code,
    extract_codes_from_text,
    validate_isbn13,
    validate_isbn10,
)


class TestValidateISBN13:
    def test_valid_isbn13(self):
        assert validate_isbn13("9784101010014") is True

    def test_invalid_isbn13_wrong_check(self):
        assert validate_isbn13("9784101010015") is False

    def test_invalid_isbn13_too_short(self):
        assert validate_isbn13("978410101") is False

    def test_invalid_isbn13_not_digits(self):
        assert validate_isbn13("978410101001X") is False


class TestValidateISBN10:
    def test_valid_isbn10(self):
        assert validate_isbn10("4101010013") is True

    def test_valid_isbn10_with_x(self):
        # ISBN-10ではチェックディジットがXになるものがある
        assert validate_isbn10("155404295X") is True

    def test_invalid_isbn10_wrong_length(self):
        assert validate_isbn10("41010100") is False


class TestNormalizeCode:
    def test_isbn13_passthrough(self):
        result = normalize_code("9784101010014")
        assert result == "9784101010014"

    def test_isbn13_with_hyphens(self):
        result = normalize_code("978-4-10-101001-4")
        assert result == "9784101010014"

    def test_isbn10_to_isbn13(self):
        result = normalize_code("4101010013")
        assert result is not None
        assert len(result) == 13
        assert result.startswith("978")

    def test_asin(self):
        result = normalize_code("B0CXYZ1234")
        assert result == "B0CXYZ1234"

    def test_asin_lowercase(self):
        result = normalize_code("b0cxyz1234")
        assert result == "B0CXYZ1234"

    def test_empty_string(self):
        result = normalize_code("")
        assert result is None

    def test_none_like(self):
        result = normalize_code("   ")
        assert result is None

    def test_long_code_extracts_isbn13(self):
        # 13桁以上で978始まりなら先頭13桁をISBNとして切出し
        result = normalize_code("978410101001499999")
        assert result == "9784101010014"

    def test_short_numeric(self):
        # 8桁以上の数字列はそのまま返す
        result = normalize_code("12345678")
        assert result == "12345678"

    def test_non_normalizable(self):
        # 正規化不能でも元文字列を返す
        result = normalize_code("ABC")
        assert result == "ABC"


class TestExtractCodesFromText:
    def test_extract_isbn13(self):
        text = "注文商品: 9784101010014 をお届けします"
        codes = extract_codes_from_text(text)
        assert "9784101010014" in codes

    def test_extract_isbn13_with_hyphens(self):
        text = "ISBN: 978-4-10-101001-4"
        codes = extract_codes_from_text(text)
        assert "9784101010014" in codes

    def test_extract_multiple_codes(self):
        text = """
        商品1: 9784101010014
        商品2: 9784062938426
        """
        codes = extract_codes_from_text(text)
        assert len(codes) >= 2

    def test_extract_asin(self):
        text = "Amazon ASIN: B0CXYZ1234 の商品"
        codes = extract_codes_from_text(text)
        assert "B0CXYZ1234" in codes

    def test_no_codes_in_text(self):
        text = "こんにちは、これは普通のテキストです。"
        codes = extract_codes_from_text(text)
        assert codes == []

    def test_deduplication(self):
        text = "9784101010014 と 9784101010014 は同じ"
        codes = extract_codes_from_text(text)
        assert codes.count("9784101010014") == 1
