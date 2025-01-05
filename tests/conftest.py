import pytest
import wmi
from lxml import etree

from noteslib.core import Session
from noteslib.db import Database
from noteslib.doc import Document, DocumentCollection
from noteslib.enums import ACLLEVEL, EMBED

DBSERVER = ''
DBPATH = '__test__.nsf'

PREFIX = 'http://www.lotus.com/dxl'
NS = {'n': PREFIX}

# Notes constants
DXLIMPORTOPTION_REPLACE_ELSE_IGNORE = 5

NOTES_EXES = ["nlnotes.exe", "notes.exe", "notes2.exe"]
WMI = wmi.WMI()
if not any(WMI.Win32_Process(Name=n) for n in NOTES_EXES):
    raise RuntimeError("A fully configured and running Notes client is needed for the tests to work")


@pytest.fixture(scope='session')
def load_notes_db():
    """Return NotesSession and NotesDatabase test objects"""
    ns = Session()
    db = ns.GetDatabase(DBSERVER, DBPATH, False)
    if not db:
        dbdir = ns.GetDbDirectory('')
        db = dbdir.CreateDatabase(DBPATH)
        assert db, 'Could not create database'
    set_acl(db)
    create_views(ns, db)
    set_title(db)
    yield ns, db
    del ns, db


@pytest.fixture
def db_with_doc0(load_notes_db):
    """Add 'doc 0' to DB and then remove it"""
    ns, db = load_notes_db
    doc = get_or_create_doc(db, [0, 0, 0])
    doc.ReplaceItemValue('Value', 'First!')
    dt = ns.CreateDateTime("Today 12:00")
    localzone = dt.LocalTime.split(" ")[-1]
    dt = ns.CreateDateTime("January 1, 2001 12:34:56 " + localzone)
    doc.ReplaceItemValue("TestDate", dt)
    dt = ns.CreateDateTime("January 1, 2001 12:34:56 GMT")
    doc.ReplaceItemValue("TestDateGMT", dt)
    if not doc.HasItem("Body2"):
        body = doc.CreateRichTextItem("Body2")
        body.EmbedObject(EMBED.ATTACHMENT, "", __file__)
    doc.Save(1, 0, 1)
    yield Database(obj=db)
    doc.Remove(True)


@pytest.fixture
def db_with_doc_cat(load_notes_db):
    """Add docs with cat to DB and then remove them"""
    _, db = load_notes_db
    docs = []
    for i in range(1, 11):
        for j in range(1, 11):
            key = ['CatTest', f'Cat1_{i:02d}', f'Cat2_{j:02d}']
            doc = get_or_create_doc(db, key)
            doc.ReplaceItemValue('Value', '-'.join(key))
            doc.Save(1, 0, 1)
            docs.append(doc)
    yield Database(obj=db)
    for doc in docs:
        doc.Remove(True)


@pytest.fixture
def temp_doc(load_notes_db):
    _, db = load_notes_db
    doc = db.CreateDocument()
    yield doc
    doc.Remove(True)


@pytest.fixture
def doc0(db_with_doc0):
    db = db_with_doc0
    vw = db.GetView("CatView")
    doc = vw.GetDocumentByKey([0, 0, 0])
    return Document(obj=doc)


@pytest.fixture
def docs0(db_with_doc0):
    db = db_with_doc0
    vw = db.GetView("CatView")
    docs = vw.GetAllDocumentsByKey([0, 0, 0])
    return DocumentCollection(obj=docs)


@pytest.fixture
def docs_cat(db_with_doc_cat):
    docs = db_with_doc_cat.AllDocuments
    return DocumentCollection(obj=docs)


@pytest.fixture
def all_docs(load_notes_db):
    _, db = load_notes_db
    return DocumentCollection(obj=db.AllDocuments)


# Auxiliary functions


def fixview(ns, vw, startcol, endcol):
    if not vw.Columns[0].IsIcon:
        db = vw.Parent
        exporter = ns.CreateDXLExporter()
        importer = ns.CreateDXLImporter()
        importer.DesignImportOption = DXLIMPORTOPTION_REPLACE_ELSE_IGNORE
        parser = etree.XMLParser(remove_blank_text=True)
        docvw = db.GetDocumentByUNID(vw.NotesURL.split('/')[-1].split('?')[0])
        tree = etree.fromstring(exporter.Export(docvw), parser)
        cols = tree.findall('.//n:column', namespaces=NS)
        cols[0].set('showaslinks', 'false')
        cols[0].set('showasicons', 'true')
        for col in cols[startcol:endcol]:
            col.set('categorized', 'true')
            col.set('twisties', 'true')
        form = etree.tostring(tree, pretty_print=True, encoding='unicode')
        importer.Import(form, db)


def get_or_create_doc(db, key):
    ns = Session()
    if not isinstance(key, list):
        key = [key]
    vw = db.GetView('CatView')
    doc = vw.GetDocumentByKey(key, True)
    if doc is None:
        doc = db.CreateDocument()
        ri = doc.ReplaceItemValue
        ri('Form', 'Test')
        lcat = []
        for i, cat in enumerate(key, 1):
            ncat = f'Category_{i}'
            lcat.append(str(cat))
            ri(ncat, cat)
        ri('Categories', '\\'.join(lcat))
        if not doc.HasItem("Body"):
            body = doc.CreateRichTextItem("Body")
            style = ns.CreateRichTextStyle()
            style.NotesFont = 1  # FONT_HELV
            body.AppendStyle(style)
            body.AppendText("Test")
        doc.Save(1, 0, 1)
    return doc


def set_title(db):
    # Set title.
    # Notes can be a little stubborn if you want to set the title programatically
    # and also add an inheritance value. Three times do the trick.
    title = 'Test DB for noteslib module'
    noteid = 'FFFF0010'
    for _ in range(3):
        doc = db.GetDocumentByID(noteid)
        doc.ReplaceItemValue('$TITLE', title + chr(10) + '#2noteslib_test')
        doc.Save(1, 0, 1)
        del doc


def create_docs(ns, db):
    # Specific sets of documents needed
    doc = get_or_create_doc(db, [0, 0, 0])
    doc.ReplaceItemValue('Value', 'First!')
    dt = ns.CreateDateTime("Today 12:00")
    localzone = dt.LocalTime.split(" ")[-1]
    dt = ns.CreateDateTime("January 1, 2001 12:34:56 " + localzone)
    doc.ReplaceItemValue("TestDate", dt)
    dt = ns.CreateDateTime("January 1, 2001 12:34:56 GMT")
    doc.ReplaceItemValue("TestDateGMT", dt)
    if not doc.HasItem("Body2"):
        body = doc.CreateRichTextItem("Body2")
        body.EmbedObject(EMBED.ATTACHMENT, "", __file__)
    doc.Save(1, 0, 1)
    vw = db.GetView("($All)")
    docs = vw.GetAllDocumentsByKey('CatTest', True)
    if docs.Count == 0:
        for i in range(1, 11):
            for j in range(1, 11):
                key = ['CatTest', f'Cat1_{i:02d}', f'Cat2_{j:02d}']
                doc = get_or_create_doc(db, key)
                doc.ReplaceItemValue('Value', '-'.join(key))
                doc.Save(1, 0, 1)


def create_views(ns, db):
    vw = db.GetView('($All)')
    if not vw:
        vw = db.CreateView("($All)", '', None, True)
        for i in range(1, 4):
            cat = f'Category_{i}'
            col = vw.CreateColumn(i, cat, cat)
            col.IsSorted = True
            col.Width = 15
        col = vw.CreateColumn(4, "Form", "Form")
        col.Width = 10
        col = vw.CreateColumn(5, "UNID", "@Text(@DocumentUniqueID)")
        col.Width = 32
    vw1 = db.GetView('CatView')
    if not vw1:
        vw1 = db.CreateView("CatView", None, vw, False)
        vw1.CreateColumn(1, " ", "IconValue")
    vw1.SelectionFormula = vw.SelectionFormula
    for i in range(0, 4):
        col = vw1.Columns[i]
        col.Width = 1
        col.Title = ''
    fixview(ns, vw1, 1, 4)
    vw2 = db.GetView('CatView2')
    if not vw2:
        vw2 = db.CreateView("CatView2", None, vw1, False)
        vw2.RemoveColumn(3)
        vw2.RemoveColumn(3)
        col = vw2.Columns[1]
        col.Formula = "Categories"
        col.Width = 32
    vw2.SelectionFormula = vw1.SelectionFormula
    fixview(ns, vw2, 1, 2)


def set_acl(db):
    acl = db.ACL
    if "[TestRole]" not in acl.Roles:
        acl.AddRole("TestRole")
    acle = acl.GetEntry('-Default-')
    if not acle:
        acle = acl.CreateACLEntry('-Default-', 6)
    acle.EnableRole("TestRole")
    acle = acl.GetEntry('John Doe/Test')
    if not acle:
        acle = acl.CreateACLEntry('John Doe/Test', 6)
        acle.Level = ACLLEVEL.AUTHOR
    acl.Save()
