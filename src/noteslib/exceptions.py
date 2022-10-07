# pylint: disable=multiple-statements
# fmt: off
class NotesLibError(Exception): pass
class SessionError(NotesLibError): pass
class DatabaseError(NotesLibError): pass
# fmt: on
