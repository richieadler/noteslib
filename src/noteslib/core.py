"""
Main classes to interact with Notes, and other useful classes available in the initial version of NotesLib
"""
from typing import Any, Dict

import win32com.client

from noteslib.exceptions import SessionError


class NotesLibObject:
    """Generic NotesLib object to incorporate methods and properties common to all objects in the library"""

    def __init__(self, *, obj=None):
        self._handle = obj

    def __eq__(self, other):
        """Two NotesLibObjects are equal if they point to the same Notes object"""
        return self.notesobj == other.notesobj

    def __getattr__(self, name):
        """Delegate to the Notes object to support all properties and methods."""
        return getattr(self._handle, name)

    @property
    def notesobj(self):
        """Returns the original Notes object"""
        return self._handle


class Session(NotesLibObject):  # pylint: disable=too-few-public-methods
    r"""
    The Session class creates an COM connection to Notes. It supports all
    the properties and methods of the LotusScript NotesSession class, using
    the same syntax.

    To create a Session object:

        s = noteslib.Session(password)

    or

        s = noteslib.Session()

    The password is optional; if you don't provide it, Notes will prompt you
    for a password.

    Example:

        >>> import noteslib
        >>> s = noteslib.Session("password")
        >>> s.NotesBuildVersion
        166
        >>> s.GetEnvironmentString("Directory", -1)
        'd:\\notes5.8\\Data'
        >>>

    Session is a Borg - multiple Session instances share status.
    You can instantiate Sessions as needed without a performance penalty
    nor errors derived of being different NotesSession objects underneath,
    and you only have to establish a password once.
    """

    _shared_state: Dict[str, Any] = {}

    def __init__(self, password=None):
        self.__dict__ = self._shared_state
        obj = self.__dict__.get("_handle", self._connect_to_notes(password))
        super().__init__(obj=obj)

    @staticmethod
    def _connect_to_notes(password=None):
        """Connect to Notes via COM."""
        try:
            obj = win32com.client.Dispatch("Lotus.NotesSession")
            if password:
                obj.Initialize(password)
            else:
                obj.Initialize()
            return obj
        except Exception as exc:
            raise SessionError() from exc
