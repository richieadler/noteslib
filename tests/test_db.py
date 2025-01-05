# ruff: noqa:
import os

import pytest
from pythoncom import com_error

from noteslib import Database, DbDirectory, DbDirectoryError

CACHE_DB = ("", "cache.ndk")

# Notes constant
ERR_SYS_FILE_NOT_FOUND = 4003


def test_db():
    db1 = Database(*CACHE_DB)
    db2 = Database(*CACHE_DB)
    db3 = Database("", "", obj=db2.notesobj)
    assert db1 == db2
    assert db2 == db3
    assert db1 == db3
    assert db1 is not db2


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
    error_code = exc_info.excepinfo[5] & 0xFFFF
    assert error_code == ERR_SYS_FILE_NOT_FOUND
