from noteslib import Session


def test_session():
    ns1 = Session()
    ns2 = Session()
    assert ns1 == ns2
