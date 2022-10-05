
class DocumentCollection:
    """Pythonic version of NotesDocumentCollection"""

    def __init__(self, nc):
        """Initialize with the passed collection"""
        self.__coll = nc

    def __len__(self):
        return self.__coll.Count

    def __getitem__(self, item):
        if not isinstance(item, int):
            raise IndexError(f"{self.__class__.__name__} indices must be integers")
        ix = (item % self.__len__()) + 1
        return self.__coll.GetNthDocument(ix)

    def __iter__(self):
        coll = self.__coll
        doc = coll.GetFirstDocument()
        while doc:
            yield doc
            doc = coll.GetNextDocument(doc)

    def __reversed__(self):
        coll = self.__coll
        doc = coll.GetLastDocument()
        while doc:
            yield doc
            doc = coll.GetPrevDocument(doc)

    def __getattr__(self, name):
        """Delegate to the Notes object to support all properties and methods."""
        return getattr(self.__coll, name)
