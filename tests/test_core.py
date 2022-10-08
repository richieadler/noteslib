from noteslib import ACL, Database, Session


def test_session():
    ns1 = Session()
    ns2 = Session()

    db1 = ns1.GetDatabase("", "cache.ndk")
    db2 = ns2.GetDatabase("", "cache.ndk")

    assert db1 == db2
    assert ns1 == ns2


def test_database():
    db = Database("", "cache.ndk")
    db2 = Database("", "cache.ndk")
    assert db == db2
    assert db is not db2


def test_acl(load_notes_db):
    _, db = load_notes_db
    acl = ACL(db.Server, db.FilePath)
    assert len(acl.entries) == 2


def test_native_properties(load_notes_db):
    ns, db = load_notes_db
    acl = ACL(db.Server, db.FilePath)
    assert "/" in ns.UserName
    assert hasattr(db.Created, "tzinfo")
    assert acl.entries[0].Level == 6
    assert acl.roles == acl.Roles
