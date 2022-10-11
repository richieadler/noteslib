from noteslib import ACL, Database, Session
from pytest_check import check_func

CACHE_DB = ("", "cache.ndk")


@check_func
def test_session():
    ns1 = Session()
    ns2 = Session()
    assert ns1 == ns2


@check_func
def test_database():
    db1 = Database(*CACHE_DB)
    db2 = Database(*CACHE_DB)
    db3 = Database("", "", obj=db2.notesobj)
    assert db1 == db2
    assert db2 == db3
    assert db1 == db3
    assert db1 is not db2


def test_acl(load_notes_db):
    _, db = load_notes_db
    acl1 = ACL(db.Server, db.FilePath)
    acl2 = ACL("", "", obj=db.ACL)
    print(acl1.entries())
    assert len(acl1.entries()) == 2
    assert acl1 == acl2


@check_func
def test_native_properties(load_notes_db):
    ns, db = load_notes_db
    acl = ACL(db.Server, db.FilePath)
    assert "/" in ns.UserName
    assert hasattr(db.Created, "tzinfo")
    assert acl.entries()[0].Level == 6
    assert acl.roles == acl.Roles

