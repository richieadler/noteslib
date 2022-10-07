"""
DocumentCollection: Wrapper object for NotesDocumentCollection, but iterable (even reversed)
and with access by index to their elements.
"""
from noteslib.doc import Document


class DocumentCollection:
    """Pythonic version of NotesDocumentCollection"""

    def __init__(self, nc):
        """Initialize with the passed collection"""
        self.__handle = nc

    def __len__(self):
        return self.__handle.Count

    def __getitem__(self, item):
        if not isinstance(item, int):
            raise IndexError(f"{self.__class__.__name__} indices must be integers")
        ix = (item % self.__handle.Count) + 1
        return Document(self.__handle.GetNthDocument(ix))

    def __iter__(self):
        coll = self.__handle
        doc = coll.GetFirstDocument()
        while doc:
            yield Document(doc)
            doc = coll.GetNextDocument(doc)

    def __reversed__(self):
        coll = self.__handle
        doc = coll.GetLastDocument()
        while doc:
            yield Document(doc)
            doc = coll.GetPrevDocument(doc)

    def __getattr__(self, name):
        """Delegate to the Notes object to support all properties and methods."""
        return getattr(self.__handle, name)
