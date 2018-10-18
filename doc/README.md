# NotesLib

NotesLib is a library of Python classes for manipulating Lotus
Notes/Domino objects via COM.

NotesLib was created by Robert Follek, and the current maintainer is Marcelo Huerta.

The NotesLib classes correspond to the standard LotusScript classes; they
support all the standard properties and methods. The NotesLib classes have
additional methods and ease-of-use features. See below the details for the
individual classes.

Classes available so far:

-   Session
-   Database
-   ACL
-   ACLEntry

## Session

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

## Database

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
NotesDatabase object via the \__handle variable.

## ACL

The ACL class encapsulates a Notes database ACL. It supports all the
properties and methods of the LotusScript NotesACL class, using the same
syntax.

Additional features:

* You can print an ACL object. It knows how to format itself reasonably.
* getAllEntries() method - Returns the ACL contents as a list of ACLEntry objects, sorted by Name.

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

## ACLEntry

The ACLEntry class encapsulates a Notes database ACL entry. It supports
all the properties and methods of the LotusScript NotesACLEntry class,
using the same syntax.

Additional features:

* You can print an ACLEntry object. It knows how to format itself reasonably.
* getName() method - Returns the entry name.
* getLevel() method - Returns the entry level.
* getRoles() method - Returns a list of entry roles, sorted alphabetically.
* getFlags() method - Returns a list of the ACLEntry flags, translated to strings.

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
