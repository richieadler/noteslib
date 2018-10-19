#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""Python classes for manipulating Lotus Notes/Domino objects via COM"""

################################################
# Copyright (c) 2001-2008 Robert I. Follek (bfollek@gmail.com)
#
# Permission is hereby granted, free of charge, to any person
# obtaining a copy of this software and associated documentation
# files (the "Software"), to deal in the Software without
# restriction, including without limitation the rights to use,
# copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the
# Software is furnished to do so, subject to the following
# conditions:
#
# The above copyright notice and this permission notice shall be
# included in all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
# EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
# OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
# NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
# HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
# WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
# FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
# OTHER DEALINGS IN THE SOFTWARE.
################################################

"""
noteslib.py is a library of Python classes for manipulating Lotus
Notes/Domino objects via COM.

The noteslib classes correspond to the standard LotusScript classes; they
support all the standard properties and methods. The noteslib classes have
additional methods and ease-of-use features. See the documentation for the
individual classes.

Classes available so far:
    Session
    Database
    ACL
    ACLEntry
"""

__version__ = "2.1.0"

import win32com.client

# -------------------------------------------------------------------

class NotesLibError(Exception): pass
class SessionError(NotesLibError): pass
class DatabaseError(NotesLibError): pass

# -------------------------------------------------------------------

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

    Error connecting to Notes via COM.
    """

    def __init__(self):
        self.__handle = None

    def __connectToNotes(self, password=None):
        """Connect to Notes via COM."""
        try:
            self.__handle = win32com.client.Dispatch("Lotus.NotesSession")
            if password:
                self.__handle.Initialize(password)
            else:
                self.__handle.Initialize()
        except:
            raise SessionError(self.__CONNECT_ERROR)

    def __call__(self, password=None):
        """Executes when an instance is invoked as a function. Singleton support."""
        if not self.__handle:
            self.__connectToNotes(password)
        return self

    def __getattr__(self, name):
        """Delegate to the Notes object to support all properties and methods."""
        return getattr(self.__handle, name)

Session = Session() # Singleton support.

# end class Session
# -------------------------------------------------------------------

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

    def __init__(self, server, dbFile, password=None):
        """Set the db handle, either from cache or via the COM connection."""
        cacheKey = ( server.lower(), dbFile.lower() )
        cachedHandle = self.__handleCache.get(cacheKey)
        if cachedHandle:
            self.__handle = cachedHandle
        else:
            try:
                s = Session(password)
                self.__handle = s.GetDatabase(server, dbFile)
                if self.__handle.IsOpen: # Make sure everything's okay.
                    self.__handleCache[cacheKey] = self.__handle # Cache the handle
            except:
                raise DatabaseError(self.__DB_ERROR % (server, dbFile))

    def __getattr__(self, name):
        """Delegate to the Notes object to support all properties and methods."""
        return getattr(self.__handle, name)

# end class Database
# -------------------------------------------------------------------

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
    ...     print (entry.getName())
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

    def __init__(self, server, dbFile, password=None):
        """Set the ACL handle, and retrieve the ACL entries."""
        self.__entries = []
        db = Database(server, dbFile, password)
        self.__handle = db.ACL
        nextEntry = self.__handle.GetFirstEntry()
        while nextEntry:
            self.__entries.append( ACLEntry(nextEntry) )
            nextEntry = self.__handle.GetNextEntry(nextEntry)
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
            s += "%s\n" % entry
        return s

# end class ACL
# -------------------------------------------------------------------

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

    __LEVELS = ["No Access", "Depositor", "Reader", "Author", "Editor", "Designer", "Manager"]

    def __init__(self, notesACLEntry):
        """The parameter is a LotusScript NotesACLEntry object."""
        self.__handle = notesACLEntry
        self.__name = notesACLEntry.Name
        self.__level = self.__LEVELS[notesACLEntry.Level]
        self.__loadRoles(notesACLEntry)
        self.__loadFlags(notesACLEntry)

    def getName(self):
        """Returns the ACLEntry Name."""
        return self.__name

    def getLevel(self):
        """Returns the ACLEntry Level, translated to a string."""
        return self.__level

    def getFlags(self):
        """Returns a list of the ACLEntry flags, translated to strings."""
        return self.__flags

    def getRoles(self):
        """Returns a list of the ACLEntry roles, sorted alphabetically."""
        return self.__roles

    def __loadFlags(self, notesACLEntry):
        """Translate the entry's flags into a list of strings."""
        self.__flags = []
        if notesACLEntry.CanCreateDocuments:
            self.__flags.append("Create Documents")
        if notesACLEntry.CanDeleteDocuments:
            self.__flags.append("Delete Documents")
        if notesACLEntry.CanCreatePersonalAgent:
            self.__flags.append("Create Personal Agents")
        if notesACLEntry.CanCreatePersonalFolder:
            self.__flags.append("Create Personal Folders/Views")
        if notesACLEntry.CanCreateSharedFolder:
            self.__flags.append("Create Shared Folders/Views")
        if notesACLEntry.CanCreateLSOrJavaAgent:
            self.__flags.append("Create LotusScript/Java Agent")
        if notesACLEntry.IsPublicReader:
            self.__flags.append("Read Public Documents")
        if notesACLEntry.IsPublicWriter:
            self.__flags.append("Write Public Documents")

    def __loadRoles(self, notesACLEntry):
        """Load the entry's roles into a sorted list."""
        roles = notesACLEntry.Roles
        if roles:
            self.__roles = sorted(roles)
        else:
            self.__roles = []

    def __lt__(self, other):
        """For sorting: compare on name."""
        return self.__name.lower() < other.__name.lower()

    def __getattr__(self, name):
        """Delegate to the Notes object to support all properties and methods."""
        return getattr(self.__handle, name)

    def __str__(self):
        """For printing"""
        s = "Name : %s\nLevel: %s\n" % ( self.getName(), self.getLevel() )
        if self.getRoles():
            for role in self.getRoles():
                s += "Role : %s\n" % role
        else:
            s += "Role : No roles\n"
        if self.getFlags():
            for flag in self.getFlags():
                s +=  "Flag : %s\n" % flag
        else:
            s += "Flag : No flags\n"
        return s

# end class ACLEntry
# -------------------------------------------------------------------

def test():
    testServer = ""
    testDB = "cache.ndk"

    print("Testing Session")
    s = Session()
    print(s.CommonUserName)

    print("Testing Database")
    db = Database(testServer, testDB)
    print(db.Created)

    print("Testing ACL")
    acl = ACL(testServer, testDB)
    print(acl.Roles)

    print("Testing the full ACL print")
    print(acl)

if __name__ == "__main__":
    test()

