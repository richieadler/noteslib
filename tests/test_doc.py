import datetime

import pendulum
import pytest
from pytest_check import check

from noteslib import Document, DocumentCollection
from noteslib.enums import DATECONV


def test_doc(doc0):
    doc = Document(obj=doc0)
    dict_doc = doc.dict()
    with check: assert dict_doc["Form"] == ["Test"]
    with check: assert dict_doc["TestDateGMT"][0] == pendulum.datetime(2001, 1, 1, 12, 34, 56, tz="GMT")
    with check: assert "$Revisions" not in doc.json(omit_special=True)


def test_get_by_index(doc0):
    doc = doc0

    with pytest.raises(KeyError):
        _ = doc["Non-existing"]
    with check: assert doc["Category_1"][0] == 0
    with check: assert doc["Body"][0] == "Test"


def test_doc_from_doccoll(load_notes_db):
    _, db = load_notes_db
    docs = DocumentCollection(obj=db.Search("Category_1 = 0", None, 0))
    doc1 = docs[0]
    doc2 = next(iter(docs))
    with check: assert doc1["Category_1"][0] == 0
    with check: assert doc2["Category_1"][0] == 0
    with check: assert doc1 == doc2


def test_doc_dates(load_notes_db, doc0):
    ns, db = load_notes_db

    # Get local Notes timezone
    dt = ns.CreateDateTime("Today 12:00")
    localzone = dt.LocalTime.split(" ")[-1]

    # Default: datetime.datetime with timezone
    retdate = doc0.get("TestDate")[0]
    with check: assert isinstance(retdate, datetime.datetime)
    with check: assert retdate == pendulum.datetime(2001, 1, 1, 12, 34, 56, tz="local")
    with check: assert doc0.get("TestDate", convert_date=DATECONV.NAIVE)[0] == datetime.datetime(2001, 1, 1, 12, 34, 56)
    with check: assert doc0.get("TestDate", convert_date=DATECONV.NATIVESTRING)[0] == "01/01/2001 12:34:56 " + localzone
    with check: assert doc0.get("TestDateGMT", convert_date="tz:Etc/GMT+1:str")[0] == "2001-01-01T11:34:56-01:00"


def test_doc_dict(doc0):
    dd = doc0.dict(convert_date="tz:GMT:str")
    with check: assert dd["TestDateGMT"][0] == "2001-01-01T12:34:56+00:00"
    with check: assert "$FILE" not in dd
    with check: assert "$Revisions" in dd


def test_len_coll(load_notes_db):
    ns, db = load_notes_db
    coll = DocumentCollection(obj=db.AllDocuments)
    assert len(coll) == coll.Count


def test_index_coll(docs0):
    coll = docs0
    with pytest.raises(IndexError):
        assert coll["a"] == ""
    assert coll[0].GetItemValue("Category_1")[0] == 0


def test_next_coll(docs0):
    coll = docs0
    doc = next(iter(coll))
    with check: assert doc.GetItemValue("Category_1")[0] == 0
    doc = next(iter(reversed(coll)))
    with check: assert doc.GetItemValue("Value")[0] == "CatTest-Cat1_10-Cat2_10"
