from noteslib import Database, Session


def test_session():
    ns1 = Session()
    ns2 = Session()

    db1 = ns1.GetDatabase("", "cache.ndk")
    db2 = ns2.GetDatabase("", "cache.ndk")

    assert ns1 == ns2


def test_database():
    db = Database("", "cache.ndk")
    db2 = Database("", "cache.ndk")
    assert db == db2
    assert db is not db2
