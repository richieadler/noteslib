"""
Exceptions for operations with Notes objects
"""

from typing import Union


class NotesLibError(Exception):
    """Generic library exception"""


class SessionError(NotesLibError):
    """Problem connecting to Notes via COM"""

    def __init__(self, message: Union[str, None] = None):
        prefix = "Error connecting to Notes via COM"
        msg = prefix + ("" if message is None else f": {message}")
        super().__init__(msg)


class DatabaseError(NotesLibError):
    """Problem connecting to a Notes database"""


class DbDirectoryError(NotesLibError):
    """Error with DbDirectory"""
