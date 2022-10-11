"""
Database related classes
"""

from typing import Any, Dict

from .core import NotesLibObject, Session
from .enums import ACLFLAGS, ACLLEVEL, ACLTYPE
from .exceptions import DatabaseError


class Database(NotesLibObject):
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
        NotesDatabase object via the _handle variable.
    """

    __DB_ERROR = r"""

    Error connecting to %s %s.

    Double-check the server and database file names, and make sure you have
    read access to the database.
    """

    __handleCache: Dict[tuple, Any] = {}

    # TODO: Wrap Database.ACL with our own ACL

    def __init__(self, server, db_path, password=None, *, obj=None):
        """Set the db handle, either from cache or via the COM connection; or use the passed NotesDatabase"""
        if obj:
            server, db_path = obj.Server, obj.FilePath
        cache_key = (server.lower(), db_path.lower())
        cached_handle = self.__handleCache.get(
            cache_key, self._get_db(cache_key, password)
        )
        super().__init__(obj=cached_handle)

    def _get_db(self, cache_key, password):
        try:
            s = Session(password)
            obj = s.GetDatabase(*cache_key)
            if obj.IsOpen:  # Make sure everything's okay.
                self.__handleCache[cache_key] = obj  # Cache the handle
            return obj
        except Exception as exc:
            raise DatabaseError(self.__DB_ERROR % cache_key) from exc


class ACL(NotesLibObject):
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

    def __init__(self, server, db_path, password=None, *, obj=None):
        """Set the ACL handle, and retrieve the ACL entries."""
        if obj is None:
            db = Database(server, db_path, password)
            handle = db.ACL
        else:
            if hasattr(obj, "Title"):  # Database
                handle = obj.ACL
            elif hasattr(obj, "GetFirstEntry"):  # ACL
                handle = obj
            else:
                raise ValueError(
                    "The object passed is neither NotesACL nor NotesDatabase"
                )
        super().__init__(obj=handle)

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
        next_entry = self._handle.GetFirstEntry()
        while next_entry:
            entry_list.append(ACLEntry(obj=next_entry))
            next_entry = self._handle.GetNextEntry(next_entry)
        entry_list.sort()
        return entry_list


class ACLEntry(NotesLibObject):
    r"""
    The ACLEntry class encapsulates a Notes database ACL entry. It supports
    all the properties and methods of the LotusScript NotesACLEntry class,
    using the same syntax.

    Additional features:

    * You can print an ACLEntry object. It knows how to format itself reasonably.

    Normally, you won't create an ACLEntry object directly. Instead, you can
    retrieve a list of ACLEntry objects from an ACL object, via its
    ``entries`` method.

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

    def __init__(self, *, obj=None):
        """The parameter is a COM NotesACLEntry object."""
        self._level = ACLLEVEL(obj.Level)
        self._type = ACLTYPE(obj.UserType)
        self._load_flags(obj)
        super().__init__(obj=obj)

    @property
    def name(self):
        """Returns the ACLEntry Name."""
        return self.notesobj.Name

    @property
    def level(self):
        """Returns the ACLEntry Level, translated to a string."""
        return str(self._level.name).title()

    @property
    def type(self):
        """Returns the ACLEntry type"""
        return str(self._type.name).title()

    @property
    def flags(self):
        """Returns a list of the ACLEntry flags, translated to strings."""
        return ", ".join(
            _.replace("_", " ").title()
            for _ in str(self._flags).split(".")[1].split("|")
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
        self._flags = ACLFLAGS(flags)

    def __lt__(self, other):
        """For sorting: compare on name."""
        return self.name.casefold() < other.name.casefold()

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
