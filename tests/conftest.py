
import pytest
import noteslib
import wmi

from lxml import etree

DBSERVER = ''
DBPATH = '__test__.nsf'

PREFIX = u'http://www.lotus.com/dxl'
NS = {'n': PREFIX}


NOTES_EXES = ["nlnotes.exe", "notes.exe", "notes2.exe"]
WMI = wmi.WMI()
if not any(WMI.Win32_Process(Name=n) for n in NOTES_EXES):
    raise RuntimeError("A fully configured and running Notes client is needed for the tests to work")


def fixview(ns, vw, startcol, endcol):
    if not vw.Columns[0].IsIcon:
        db = vw.Parent
        exporter = ns.CreateDXLExporter()
        importer = ns.CreateDXLImporter()
        importer.DesignImportOption = 5 # DXLIMPORTOPTION_REPLACE_ELSE_IGNORE
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
        doc.Save(1,0,1)
    return doc


@pytest.fixture(scope='session')
def load_notes_db():
    ns = noteslib.Session()
    db = ns.GetDatabase(DBSERVER, DBPATH, False)
    if not db:
        dbdir = ns.GetDbDirectory('')
        db = dbdir.CreateDatabase(DBPATH)
        assert db, 'Could not create database'
    acl = db.ACL
    acle = acl.GetEntry('-Default-')
    if not acle:
        acl.CreateACLEntry('-Default-', 6)
        acl.Save()
    vw = db.GetView('($All)')
    if not vw:
        vw = db.CreateView("($All)", '', None, True)
        for i in range(1,4):
            cat = f'Category_{i}'
            col = vw.CreateColumn(i, cat, cat)
            col.IsSorted = True
            col.Width = 15
        col = vw.CreateColumn(4, "Form", "Form")
        col.Width = 10
        col = vw.CreateColumn(5, "UNID", "@Text(@DocumentUniqueID)")
        col.Width = 32
    vwC = db.GetView('CatView')
    if not vwC:
        vwC = db.CreateView("CatView", None, vw, False)
        col = vwC.CreateColumn(1, " ", "IconValue")
    vwC.SelectionFormula = vw.SelectionFormula
    for i in range(0,4):
        col = vwC.Columns[i]
        col.Width = 1
        col.Title = ''
    fixview(ns, vwC, 1, 4)
    vwC2 = db.GetView('CatView2')
    if not vwC2:
        vwC2 = db.CreateView("CatView2", None, vwC, False)
        vwC2.RemoveColumn(3)
        vwC2.RemoveColumn(3)
        col = vwC2.Columns[1]
        col.Formula = "Categories"
        col.Width = 32
    vwC2.SelectionFormula = vwC.SelectionFormula
    fixview(ns, vwC2, 1, 2)

    # Specific sets of documents needed
    doc = get_or_create_doc(db, [0, 0, 0])
    doc.ReplaceItemValue('Value', 'First!')
    doc.Save(1,0,1)
    docs = vw.GetAllDocumentsByKey('CatTest', True)
    if docs.Count == 0:
        for i in range(1,11):
            for j in range(1,11):
                key = ['CatTest', f'Cat1_{i:02d}', f'Cat2_{j:02d}']
                doc = get_or_create_doc(db, key)
                doc.ReplaceItemValue('Value', '-'.join(key))
                doc.Save(1, 0, 1)

    # Title
    title = 'Test DB for noteslib module'
    doc = db.GetDocumentByID('FFFF0010')
    doc.ReplaceItemValue('$TITLE', title)
    doc.Save(1, 0, 1)
    del doc
    doc = db.GetDocumentByID('FFFF0010')
    doc.ReplaceItemValue('$TITLE', title)
    doc.Save(1, 0, 1)
    del doc
    doc = db.GetDocumentByID('FFFF0010')
    doc.ReplaceItemValue('$TITLE',  title + chr(10) + '#2noteslib_test')
    doc.Save(1, 0, 1)

    yield ns, db

    del ns, db


@pytest.fixture(scope='function')
def temp_notes_doc(load_notes_db):
    ns, db = load_notes_db
    doc = db.CreateDocument()
    yield doc
    doc.Remove(True)