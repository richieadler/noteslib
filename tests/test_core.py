from pytest_check import check_func

from noteslib import Session


@check_func
def test_session():
    ns1 = Session()
    ns2 = Session()
    assert ns1 == ns2
