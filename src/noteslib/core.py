import win32com.client

from noteslib.enums import ACLLEVEL
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

    Session is a singleton - multiple Session variables share one Session
    object. You can instantiate Sessions as needed without a performance
    penalty, and you only have to establish a password once. Example:

        >>> a = noteslib.Session(password)
        >>> id(a)
        8429868
        >>> b = noteslib.Session()
        >>> id(b)
        8429868
    """
    ################################################
    # SINGLETON - Implementation Details
    #
    # 1) The __call__ method in the Session class ensures that function-style
    # calls of a Session instance return the instance.
    #
    # 2) The line "Session = Session()" that immediately follows the Session class definition
    # creates an instance of the Session class and rebinds the name "Session" to it.
    #
    # With these pieces in place, any assignment like "s = Session()" returns the same
    # Session instance. This gives us the singleton we want.
    #
    # The attempt to connect to Notes is in Session.__call__ rather than Session.__init__
    # so that we don't try to connect when the "Session = Session()" line executes.
    # Otherwise, "import noteslib" might try to connect, fail, and raise an exception.
    ################################################

    __CONNECT_ERROR = r"""

    Error connecting to Notes via COM:
    """

    def __init__(self):
        self.__handle = None

    def __connect_to_notes(self, password=None):
        """Connect to Notes via COM."""
        try:
            self.__handle = win32com.client.Dispatch("Lotus.NotesSession")
            if password:
                self.__handle.Initialize(password)
            else:
                self.__handle.Initialize()
        except Exception as exc:
            raise SessionError(self.__CONNECT_ERROR) from exc

    def __call__(self, password=None):
        """Executes when an instance is invoked as a function. Singleton support."""
        if not self.__handle:
            self.__connect_to_notes(password)
        return self

    def __getattr__(self, name):
        """Delegate to the Notes object to support all properties and methods."""
        return getattr(self.__handle, name)


Session = Session()  # Singleton support.


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
        >>> id(a._Database__handle)
        15286172
        >>> b = noteslib.Database("NYNotes1", "ACLTest.nsf")
        >>> id(b)
        15270044
        >>> id(b._Database__handle)
        15286172

        a and b are different objects, but they share the same internal
        NotesDatabase object via the __handle variable.
    """

    __DB_ERROR = r"""

    Error connecting to %s %s.

    Double-check the server and database file names, and make sure you have
    read access to the database.
    """

    __handleCache = {}

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
        return self.__handle == other.__handle

    def __getattr__(self, name):
        """Delegate to the Notes object to support all properties and methods."""
        return getattr(self.__handle, name)


class ACL:
    r"""
    The ACL class encapsulates a Notes database ACL. It supports all the
    properties and methods of the LotusScript NotesACL class, using the same
    syntax.

    Additional features:
    * You can print an ACL object. It knows how to format itself reasonably.
    * getAllEntries() method - Returns the ACL contents as a list of ACLEntry
        objects, sorted by Name.

    You don't have to create Session or Database objects first. An ACL object
    creates its own Session and Database objects automatically.

    To create an ACL object:

        acl = noteslib.ACL(server, database_file, password)

    or

        acl = noteslib.ACL(server, database_file)

    Example:

        >>> import noteslib
        >>> acl = noteslib.ACL("NYNotes1", "ACLTest.nsf", "password")
        >>> for entry in acl.getAllEntries():
        ...     print (entry.name())
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

    def __init__(self, server, db_path, password=None):
        """Set the ACL handle, and retrieve the ACL entries."""
        self.__entries = []
        db = Database(server, db_path, password)
        self.__handle = db.ACL
        next_entry = self.__handle.GetFirstEntry()
        while next_entry:
            self.__entries.append(ACLEntry(next_entry))
            next_entry = self.__handle.GetNextEntry(next_entry)
        self.__entries.sort()

    def getAllEntries(self):
        """Returns a list of noteslib ACLEntry objects, sorted by Name."""
        return self.__entries

    def __getattr__(self, name):
        """Delegate to the Notes object to support all properties and methods."""
        return getattr(self.__handle, name)

    def __str__(self):
        """For printing"""
        s = ""
        for entry in self.getAllEntries():
            s += f"{entry}\n"
        return s


class ACLEntry:
    r"""
    The ACLEntry class encapsulates a Notes database ACL entry. It supports
    all the properties and methods of the LotusScript NotesACLEntry class,
    using the same syntax.

    Additional features:
    * You can print an ACLEntry object. It knows how to format itself reasonably.
    * getName() method - Returns the entry name.
    * getLevel() method - Returns the entry level.
    * getRoles() method - Returns a list of entry roles, sorted alphabetically.
    * getFlags() method - Returns a list of the ACLEntry flags, translated to
        strings.
    These methods avoid the obvious names, e.g. getName() instead of name(),
    to avoid conflict with the existing NotesACLEntry properties.

    Normally, you won't create an ACLEntry object directly. Instead, you can
    retrieve a list of ACLEntry objects from an ACL object, via its
    getAllEntries() method.

    Example:

        >>> import noteslib
        >>> acl = noteslib.ACL("NYNotes1", "ACLTest.nsf", "password")
        >>> print (acl.getAllEntries()[3])
        Name : bob
        Level: Manager
        Role : [Role1]
        Role : [Role2]
        Role : [Role3]
        Flag : Create Documents
        Flag : Delete Documents
        Flag : Create Personal Agents
        Flag : Create Personal Folders/Views
        Flag : Create Shared Folders/Views
        Flag : Create LotusScript/Java Agent
        Flag : Read Public Documents
        Flag : Write Public Documents
    """

    def __init__(self, notes_acl_entry):
        """The parameter is a LotusScript NotesACLEntry object."""
        self.__handle = notes_acl_entry
        self.__level = ACLLEVEL(notes_acl_entry.Level)
        self._load_flags(notes_acl_entry)

    @property
    def name(self):
        """Returns the ACLEntry Name."""
        return self.__handle.Name

    @property
    def level(self):
        """Returns the ACLEntry Level, translated to a string."""
        return str(self.__level.name).title()

    def getFlags(self):
        """Returns a list of the ACLEntry flags, translated to strings."""
        return self.__flags

    @property
    def roles(self):
        """Returns a list of the ACLEntry roles, sorted alphabetically."""
        # return self.__roles
        return list(self.__handle.Roles)

    def _load_flags(self, acl_entry):
        """Translate the entry's flags into a list of strings."""
        self.__flags = []
        if acl_entry.CanCreateDocuments:
            self.__flags.append("Create Documents")
        if acl_entry.CanDeleteDocuments:
            self.__flags.append("Delete Documents")
        if acl_entry.CanCreatePersonalAgent:
            self.__flags.append("Create Personal Agents")
        if acl_entry.CanCreatePersonalFolder:
            self.__flags.append("Create Personal Folders/Views")
        if acl_entry.CanCreateSharedFolder:
            self.__flags.append("Create Shared Folders/Views")
        if acl_entry.CanCreateLSOrJavaAgent:
            self.__flags.append("Create LotusScript/Java Agent")
        if acl_entry.IsPublicReader:
            self.__flags.append("Read Public Documents")
        if acl_entry.IsPublicWriter:
            self.__flags.append("Write Public Documents")

    def __lt__(self, other):
        """For sorting: compare on name."""
        return self.name.casefold() < other.name.casefold()

    def __getattr__(self, name):
        """Delegate to the Notes object to support all properties and methods."""
        return getattr(self.__handle, name)

    def __str__(self):
        """For printing"""
        s = f"Name : {self.name}\nLevel: {self.level}\n"
        # if self.roles:
        #     for role in self.roles:
        #         s += f"Role : {role}\n"
        # else:
        #     s += "Role : No roles\n"
        s += f"Roles: {self.roles}\n"
        if self.getFlags():
            for flag in self.getFlags():
                s += f"Flag : {flag}\n"
        else:
            s += "Flag : No flags\n"
        return s
