import datetime

import pendulum
import pytest
from pytest_check import check_func

from noteslib import Document, DocumentCollection
from noteslib.enums import DATECONV


@check_func
def test_doc(doc0):
    doc = Document(obj=doc0)
    dict_doc = doc.dict()
    assert dict_doc["Form"] == ["Test"]
    assert dict_doc["TestDate"] == ['2001-01-01T15:34:56+00:00']
    assert "$Revisions" not in doc.json(omit_special=True)


@check_func
def test_get_by_index(doc0):
    doc = doc0

    with pytest.raises(KeyError):
        _ = doc["Non-existing"]
    assert doc["Category_1"][0] == 0
    assert doc["Body"][0] == "Test"


def test_doc_from_doccoll(load_notes_db):
    _, db = load_notes_db
    docs = DocumentCollection(obj=db.Search("Category_1 = 0", None, 0))
    doc1 = docs[0]
    doc2 = next(iter(docs))
    assert doc1["Category_1"][0] == 0
    assert doc2["Category_1"][0] == 0
    assert doc1 == doc2


@check_func
def test_doc_dates(load_notes_db, doc0):
    ns, db = load_notes_db

    # Get local Notes timezone
    dt = ns.CreateDateTime("Today 12:00")
    localzone = dt.LocalTime.split(" ")[-1]

    # Default: datetime.datetime with timezone
    retdate = doc0.get("TestDate")[0]
    assert isinstance(retdate, datetime.datetime)
    assert retdate == pendulum.datetime(2001, 1, 1, 12, 34, 56, tz="local")
    assert doc0.get("TestDate", convert_date=DATECONV.NAIVE)[0] == datetime.datetime(2001, 1, 1, 12, 34, 56)
    assert doc0.get("TestDate", convert_date=DATECONV.NATIVESTRING)[0] == "01/01/2001 12:34:56 " + localzone
    assert doc0.get("TestDateGMT", convert_date="tz:Etc/GMT+1:str")[0] == "2001-01-01T11:34:56-01:00"


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
    assert doc.GetItemValue("Category_1")[0] == 0
    doc = next(iter(reversed(coll)))
    assert doc.GetItemValue("Value")[0] == "CatTest-Cat1_10-Cat2_10"
