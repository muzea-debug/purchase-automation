import pytest
from pathlib import Path
from domain.services.발주_코드_서비스 import 발주_코드_서비스


def test_파싱_정상():
    prefix, 번호 = 발주_코드_서비스.파싱("NF(694)")
    assert prefix == "NF" and 번호 == 694


def test_파싱_단자리():
    prefix, 번호 = 발주_코드_서비스.파싱("EA(8)")
    assert prefix == "EA" and 번호 == 8


def test_파싱_잘못된_형식():
    with pytest.raises(ValueError):
        발주_코드_서비스.파싱("NF694")


def test_포맷():
    assert 발주_코드_서비스.포맷("NF", 695) == "NF(695)"


def test_유효성_검사_정상():
    assert 발주_코드_서비스.유효성_검사("NF(695)") is True


def test_유효성_검사_실패():
    assert 발주_코드_서비스.유효성_검사("NF695") is False
    assert 발주_코드_서비스.유효성_검사("") is False


def test_다음_번호_로컬(tmp_path):
    정기발주 = tmp_path / "정기발주"
    정기발주.mkdir()
    (정기발주 / "NF(690).xlsx").touch()
    (정기발주 / "NF(694).xlsx").touch()
    (정기발주 / "EA(8).xlsx").touch()
    assert 발주_코드_서비스.다음_번호_로컬(tmp_path, "NF") == 695


def test_다음_번호_로컬_없으면_1(tmp_path):
    (tmp_path / "정기발주").mkdir()
    assert 발주_코드_서비스.다음_번호_로컬(tmp_path, "NF") == 1
