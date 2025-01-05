# ruff: noqa: E701
import os

import pytest
from pytest_check import check
from pythoncom import com_error

from noteslib.db import Database, DbDirectory
from noteslib.doc import Document
from noteslib.exceptions import DbDirectoryError

CACHE_DB = ("", "cache.ndk")

# Notes constant
ERR_SYS_FILE_NOT_FOUND = 4003


def test_db():
    db1 = Database(*CACHE_DB)
    db2 = Database(*CACHE_DB)
    db3 = Database("", "", obj=db2.notesobj)
    # fmt: off
    with check: assert db1 == db2
    with check: assert db2 == db3
    with check: assert db1 == db3
    with check: assert db1 is not db2
    # fmt: on


def test_db_by_index(db_with_doc0, doc0):
    # Indexing a Database by unid or by noteid should return the corresponding Document
    db = db_with_doc0
    unid = doc0.UniversalID
    noteid = doc0.NoteID
    # fmt: off
    with check: assert (docunid := db[unid]) == doc0
    with check: assert (isinstance(docunid, Document))
    with check: assert db[noteid] == doc0
    with pytest.raises(KeyError):
        assert db["deadbeef"]
    with pytest.raises(KeyError):
        assert db["12345678901234567890123456789012"]
    # fmt: on


def test_dbdir():
    dbdir = DbDirectory("")
    list_db = list(filter(lambda db: (os.path.basename(db.FilePath) == "names.nsf"), dbdir.databases()))
    assert len(list_db) == 1
    with pytest.raises(DbDirectoryError):
        DbDirectory("", obj=False)


def test_dbdir_open():
    dbdir = DbDirectory("")
    db = dbdir.OpenDatabase("names.nsf")
    assert db.IsOpen
    with pytest.raises(com_error) as exc_info:
        dbdir.OpenDatabase("this_database_doesnt_exist")
    excepinfo = exc_info.value.args[2]
    error_code = excepinfo[5] & 0xFFFF
    assert error_code == ERR_SYS_FILE_NOT_FOUND
