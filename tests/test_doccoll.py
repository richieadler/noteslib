import pytest

from noteslib import DocumentCollection


def test_len(load_notes_db):
    ns, db = load_notes_db
    coll = DocumentCollection(db.AllDocuments)
    assert len(coll) == coll.Count


def test_index(load_notes_db):
    ns, db = load_notes_db
    coll = DocumentCollection(db.AllDocuments)
    with pytest.raises(IndexError):
        assert coll["a"] == ""
    assert coll[0].GetItemValue("Category_1")[0] == 0


def test_next(load_notes_db):
    ns, db = load_notes_db
    coll = DocumentCollection(db.AllDocuments)
    doc = next(iter(coll))
    assert doc.GetItemValue("Category_1")[0] == 0
    doc = next(iter(reversed(coll)))
    assert doc.GetItemValue("Value")[0] == "CatTest-Cat1_10-Cat2_10"
