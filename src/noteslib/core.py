"""
Main classes to interact with Notes, and other useful classes available in the initial version of NotesLib
"""
import sys
from typing import Any, Dict

import win32com.client

from noteslib.enums import ACLFLAGS, ACLLEVEL, ACLTYPE
from noteslib.exceptions import DatabaseError, SessionError


class Session:
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
        if self.__dict__.get("_handle") is None:
            self._connect_to_notes(password)

    def _connect_to_notes(self, password=None):
        """Connect to Notes via COM."""
        try:
            self._handle = win32com.client.Dispatch("Lotus.NotesSession")
            if password:
                self._handle.Initialize(password)
            else:
                self._handle.Initialize()
        except Exception as exc:
            raise SessionError() from exc

    def __eq__(self, other):
        return self.notesobj == other.notesobj

    def __getattr__(self, name):
        """Delegate to the Notes object to support all properties and methods."""
        return getattr(self._handle, name)

    @property
    def notesobj(self):
        """Returns the original Notes object"""
        return self._handle


class Database:
    r"""
    The Database class creates an COM connection to a Notes database. It
    supports all the properties and methods of the LotusScript NotesDatabase
    class, using the same syntax.

    You don't have to create a Session first. A Database object creates its own
    Session automatically.

    To create a Database object:

        db = noteslib.Database(server, database_file, password)

    or

        db = noteslib.Database(server, database_file)

    Example:

        >>> import noteslib
        >>> db = noteslib.Database("NYNotes1", "ACLTest.nsf", "password")
        >>> db.Created
        pywintypes.datetime(2001, 6, 30, 11, 12, 40, tzinfo=TimeZoneInfo('GMT Standard Time', True))

    Multiple Database objects created for the same database are unique objects,
    but they share the same handle to the underlying NotesDatabase object.
    You can instantiate Database objects as needed without a performance
    penalty. Example:

        >>> a = noteslib.Database("NYNotes1", "ACLTest.nsf", "password")
        >>> id(a)
        15281724
        >>> id(a.notesobj)
        15286172
        >>> b = noteslib.Database("NYNotes1", "ACLTest.nsf")
        >>> id(b)
        15270044
        >>> id(b.notesobj)
        15286172

        a and b are different objects, but they share the same internal
        NotesDatabase object via the __handle variable.
    """

    __DB_ERROR = r"""

    Error connecting to %s %s.

    Double-check the server and database file names, and make sure you have
    read access to the database.
    """

    __handleCache: Dict[tuple, Any] = {}

    # TODO: Wrap Database.ACL with our own ACL

    def __init__(self, server, db_path, password=None):
        """Set the db handle, either from cache or via the COM connection."""
        cache_key = (server.lower(), db_path.lower())
        cached_handle = self.__handleCache.get(cache_key)
        if cached_handle:
            self.__handle = cached_handle
        else:
            try:
                s = Session(password)
                self.__handle = s.GetDatabase(server, db_path)
                if self.__handle.IsOpen:  # Make sure everything's okay.
                    self.__handleCache[cache_key] = self.__handle  # Cache the handle
            except Exception as exc:
                raise DatabaseError(self.__DB_ERROR % (server, db_path)) from exc

    def __eq__(self, other):
        """Two databases are equal if they point to the same NotesDatabase object"""
        return self.notesobj == other.notesobj

    def __getattr__(self, name):
        """Delegate to the Notes object to support all properties and methods."""
        return getattr(self.__handle, name)

    @property
    def notesobj(self):
        """Returns the original Notes object"""
        return self.__handle


class ACL:
    r"""
    The ACL class encapsulates a Notes database ACL. It supports all the
    properties and methods of the LotusScript NotesACL class, using the same
    syntax.

    Additional features:
    * You can print an ACL object. It knows how to format itself reasonably.

    You don't have to create Session or Database objects first. An ACL object
    creates its own Session and Database objects automatically.

    To create an ACL object:

        acl = noteslib.ACL(server, database_file, password)

    or

        acl = noteslib.ACL(server, database_file)

    Example:

        >>> import noteslib
        >>> acl = noteslib.ACL("NYNotes1", "ACLTest.nsf", "password")
        >>> for entry in acl.entries:
        ...     print (entry.name)
        ...
        -Default-
        Alice Author
        Anonymous
        bob
        Dave Depositor
        Donna Designer
        LocalDomainServers
        OtherDomainServers
        Randy Reader
    """

    # TODO: Allow initialization with existing NotesACL

    def __init__(self, server, db_path, password=None):
        """Set the ACL handle, and retrieve the ACL entries."""
        db = Database(server, db_path, password)
        self.__handle = db.ACL

    def __getattr__(self, name):
        """Delegate to the Notes object to support all properties and methods."""
        return getattr(self.__handle, name)

    def __str__(self):
        """For printing"""
        s = ""
        for entry in self.entries:
            s += f"{entry}\n"
        return s

    @property
    def entries(self) -> list:
        """Returns a sorted list of ACLEntry objects based on the NotesACLEntry objects from the original"""
        entry_list = []
        next_entry = self.__handle.GetFirstEntry()
        while next_entry:
            entry_list.append(ACLEntry(next_entry))
            next_entry = self.__handle.GetNextEntry(next_entry)
        entry_list.sort()
        return entry_list


class ACLEntry:
    r"""
    The ACLEntry class encapsulates a Notes database ACL entry. It supports
    all the properties and methods of the LotusScript NotesACLEntry class,
    using the same syntax.

    Additional features:
    * You can print an ACLEntry object. It knows how to format itself reasonably.

    Normally, you won't create an ACLEntry object directly. Instead, you can
    retrieve a list of ACLEntry objects from an ACL object, via its
    `entries` property.

    Example:

        >>> import noteslib
        >>> acl = noteslib.ACL("NYNotes1", "ACLTest.nsf", "password")
        >>> print (acl.entries[3])
        Name : bob
        Level: Manager
        Roles: [Role1], [Role2], [Role3]
        Flags: Create Documents, Delete Documents, Create Personal Agents, Create Personal Folders And Views,
          Create Shared Folders And Views, Create Agent, Read Public Documents, Write Public Documents
    """

    def __init__(self, notes_acl_entry):
        """The parameter is a COM NotesACLEntry object."""
        self.__handle = notes_acl_entry
        self.__level = ACLLEVEL(notes_acl_entry.Level)
        self.__type = ACLTYPE(notes_acl_entry.UserType)
        self._load_flags(notes_acl_entry)

    @property
    def name(self):
        """Returns the ACLEntry Name."""
        return self.__handle.Name

    @property
    def level(self):
        """Returns the ACLEntry Level, translated to a string."""
        return str(self.__level.name).title()

    @property
    def type(self):
        """Returns the ACLEntry type"""
        return str(self.__type.name).title()

    @property
    def flags(self):
        """Returns a list of the ACLEntry flags, translated to strings."""
        return ", ".join(
            _.replace("_", " ").title()
            for _ in str(self.__flags).split(".")[1].split("|")
        )

    @property
    def roles(self):
        """Returns a list of the ACLEntry roles."""
        return ", ".join(self.__handle.Roles)

    def _load_flags(self, acl_entry):
        """Translate the entry's flags into a list of strings."""
        flags = 0
        possible = dict(
            CanCreateDocuments=ACLFLAGS.CREATE_DOCUMENTS,
            CanDeleteDocuments=ACLFLAGS.DELETE_DOCUMENTS,
            CanCreatePersonalAgent=ACLFLAGS.CREATE_PRIV_AGENTS,
            CanCreatePersonalFolder=ACLFLAGS.CREATE_PRIV_FOLDERS_VIEWS,
            CanCreateSharedFolder=ACLFLAGS.CREATE_SHARED_FOLDERS_VIEWS,
            CanCreateLSOrJavaAgent=ACLFLAGS.CREATE_SCRIPT_AGENTS,
            IsPublicReader=ACLFLAGS.READ_PUBLIC_DOCUMENTS,
            IsPublicWriter=ACLFLAGS.WRITE_PUBLIC_DOCUMENTS,
        )
        for key, value in possible.items():
            if getattr(acl_entry, key):
                flags += value
        self.__flags = ACLFLAGS(flags)

    def __lt__(self, other):
        """For sorting: compare on name."""
        return self.name.casefold() < other.name.casefold()

    def __getattr__(self, name):
        """Delegate to the Notes object to support all properties and methods."""
        return getattr(self.__handle, name)

    def __str__(self):
        """For printing"""
        s = [
            f"Name : {self.name}",
            f"Type : {self.type}",
            f"Level: {self.level}",
            f"Roles: {self.roles}",
            f"Flags: {self.flags}",
        ]
        return "\n".join(s) + "\n"
