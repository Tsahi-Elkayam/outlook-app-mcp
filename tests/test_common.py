from datetime import timedelta

import pytest

from outlook_app_mcp.common import (
    OutlookError,
    attachment_destination,
    check_addresses,
    check_ids,
    parse_when,
    pick_attachment,
    safe_filename,
    to_html,
)


def test_parse_when_forms():
    assert parse_when("2026-11-10").hour == 0
    assert parse_when("2026-11-10T14:30").minute == 30
    assert parse_when("2026-11-10 14:30").hour == 14
    assert parse_when("2026-11-10T14:30+00:00").utcoffset() == timedelta(0)
    assert parse_when("2026-11-10").tzinfo is not None
    with pytest.raises(OutlookError, match="ISO date"):
        parse_when("next tuesday")

def test_to_html_direction_and_escaping():
    h = to_html("שלום רינה,\nשורה שנייה\n\nHello <b>x</b>")
    assert h.count("<div dir=\"rtl\"") == 1 and h.count("<div dir=\"ltr\">") == 1
    assert "שלום רינה,<br>שורה שנייה" in h
    assert "&lt;b&gt;" in h and "<b>" not in h
    # first strong character decides: "avi-c@ צריך" starts Latin
    assert to_html("avi-c@ צריך").startswith('<div dir="ltr">')
    assert to_html("1. נתיב").startswith('<div dir="rtl"')

def test_check_addresses():
    assert check_addresses([" a@b.co "], "to") == ["a@b.co"]
    assert check_addresses(None, "to") == []
    for bad in ["no-at", "a@b", "a b@c.d", "a@b.c,d@e.f"]:
        with pytest.raises(OutlookError):
            check_addresses([bad], "to")


def test_check_ids_keeps_opaque_text():
    long_id = "00000000A1B2" * 12
    assert check_ids([long_id, " 5 "]) == [long_id, "5"]
    assert check_ids("x") == ["x"]
    for bad in ([], ["", "x"]):
        with pytest.raises(OutlookError):
            check_ids(bad)


def test_safe_filename_for_both_systems():
    assert safe_filename("\u200e\u2068ממצאים\u2069.docx") == "ממצאים.docx"
    assert safe_filename('a/b\\c:d*e?f"g<h>i|j.txt') == "a_b_c_d_e_f_g_h_i_j.txt"
    assert safe_filename("CON.txt") == "_CON.txt" and safe_filename("report. ") == "report"
    assert safe_filename("\u200e") == "attachment" and safe_filename("..hidden") == "hidden"


def test_attachment_destination_never_overwrites(tmp_path):
    (tmp_path / "a.pdf").write_bytes(b"x")
    assert attachment_destination(str(tmp_path), "a.pdf").name == "a (1).pdf"
    with pytest.raises(OutlookError, match="not an existing folder"):
        attachment_destination(str(tmp_path / "nope"), "a.pdf")


def test_pick_attachment():
    atts = [{"index": 1, "name": "a.pdf"}, {"index": 2, "name": "b"}, {"index": 3, "name": "b"}]
    assert pick_attachment(atts, "1", "m")["name"] == "a.pdf"
    assert pick_attachment(atts, 2, "m")["index"] == 2
    with pytest.raises(OutlookError, match="Several"):
        pick_attachment(atts, "b", "m")
    with pytest.raises(OutlookError, match="Attachments: 1: a.pdf"):
        pick_attachment(atts, "zz", "m")
