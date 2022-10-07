import pytest
from pytest_check import check_func


@check_func
def test_get_by_index(doc0):
    doc = doc0
    body = doc.CreateRichTextItem("Body")
    body.AppendText("Hello")

    with pytest.raises(KeyError):
        _ = doc["Non-existing"]
    assert doc["Category_1"][0] == 0
    assert doc["Body"] == "Hello"


def test_doc_from_doccoll(all_docs):
    doc1 = all_docs[0]
    doc2 = next(iter(all_docs))
    assert doc1["Category_1"][0] == 0
    assert doc2["Category_1"][0] == 0
    assert doc1 == doc2
