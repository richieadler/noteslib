from .enums import ITEMTYPE


class Document:
    """Wrapper for NotesDocument with "pythonic" syntax"""

    def __init__(self, doc):
        self.__handle = doc

    def __eq__(self, other):
        docself = self.__handle
        docother = other.__handle
        return (
            docself.ParentDatabase.ReplicaID == docother.ParentDatabase.ReplicaID
            and docself.UniversalID == docother.UniversalID
        )

    def __getitem__(self, name):
        """Get item value by name, as index"""
        # TODO: Date conversions
        # TODO: Allow a "Compact" mode (return a scalar if there is only 1 value in the item.Values)
        # TODO: Rich text conversions
        item = self.__handle.GetFirstItem(name)
        if item is None:
            # TODO: Different return modes for non-existing items
            raise KeyError(repr(name))
        return self._get(item)

    def __getattr__(self, name):
        """Delegate to the Notes object to support all properties and methods."""
        return getattr(self.__handle, name)

    @staticmethod
    def _get(item):
        """Return value according to type and other considerations"""

        if item.Type == ITEMTYPE.RICHTEXT:
            return item.Text
        return item.Values
