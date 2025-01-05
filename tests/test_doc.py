# ruff: noqa: E701

import datetime

import pytest
from pytest_check import check

from noteslib import Document, DocumentCollection, Session
from noteslib.enums import DATECONV
from tests.conftest import docs_cat


def test_doc(doc0):
    doc = Document(obj=doc0)
    dict_doc = doc.asdict()
    # fmt: off
    with check: assert dict_doc["Form"] == ["Test"]
    with check: assert dict_doc["TestDateGMT"][0] == datetime.datetime(2001, 1, 1, 12, 34, 56, tzinfo=datetime.timezone.utc)
    with check: assert "$Revisions" not in doc.json(omit_special=True)
    # fmt: on


def test_get_by_index(doc0):
    doc = doc0
    with pytest.raises(KeyError):
        _ = doc["Non-existing"]
    # fmt: off
    with check: assert doc["Category_1"][0] == 0
    with check: assert doc["Body"][0] == "Test"
    # fmt: on


def test_doc_from_doccoll(db_with_doc0):
    db = db_with_doc0
    docs = DocumentCollection(obj=db.Search("Category_1 = 0", None, 0))
    doc1 = docs[0]
    doc2 = next(iter(docs))
    # fmt: off
    with check: assert doc1["Category_1"][0] == 0
    with check: assert doc2["Category_1"][0] == 0
    with check: assert doc1 == doc2
    # fmt: on


def test_doc_dates(doc0):
    ns = Session()

    # Get local Notes timezone
    dt = ns.CreateDateTime("Today 12:00")
    localzone = dt.LocalTime.split(" ")[-1]
    pylocalzone = datetime.datetime.now().astimezone().tzinfo

    # Default: datetime.datetime with timezone
    retdate = doc0.get("TestDate")[0]
    # fmt: off
    with check: assert isinstance(retdate, datetime.datetime)
    with check: assert retdate == datetime.datetime(2001, 1, 1, 12, 34, 56, tzinfo=pylocalzone)
    with check: assert doc0.get("TestDate", convert_date=DATECONV.NATIVESTRING)[0] == "01/01/2001 12:34:56 " + localzone
    with check: assert doc0.get("TestDateGMT", convert_date="tz:Etc/GMT+1:str")[0] == "2001-01-01T11:34:56-01:00"
    # fmt: on


def test_doc_dict(doc0):
    dd = doc0.asdict(convert_date="tz:GMT:str")
    # fmt: off
    with check: assert dd["TestDateGMT"][0] == "2001-01-01T12:34:56+00:00"
    with check: assert "$FILE" not in dd
    with check: assert "$Revisions" in dd
    # fmt: on


def test_len_coll(load_notes_db):
    ns, db = load_notes_db
    coll = DocumentCollection(obj=db.AllDocuments)
    assert len(coll) == coll.Count


def test_index_coll(docs0):
    coll = docs0
    with pytest.raises(IndexError):
        assert coll["a"] == ""
    assert coll[0].GetItemValue("Category_1")[0] == 0


def test_next_coll(docs_cat):
    coll = docs_cat
    doc = next(iter(coll))
    # fmt: off
    with check: assert doc.GetItemValue("Category_1")[0] == "CatTest"
    doc = next(iter(reversed(coll)))
    with check: assert doc.GetItemValue("Value")[0] == "CatTest-Cat1_10-Cat2_10"
    # fmt: on
