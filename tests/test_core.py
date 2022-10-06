from noteslib import Database


def test_database():
    db = Database("", "cache.ndk")
    db2 = Database("", "cache.ndk")
    assert db == db2
