"""
Enumerations to represent Notes "magic numbers" or assorted values which are logically related
"""

import enum


class ACLFLAGS(enum.IntFlag):
    """ACL flags for an entry"""

    CREATE_DOCUMENTS = 1
    DELETE_DOCUMENTS = 2
    CREATE_PERSONAL_AGENTS = 4
    CREATE_PRIV_AGENTS = 4
    CREATE_PERSONAL_FOLDERS_AND_VIEWS = 8
    CREATE_PRIV_FOLDERS_VIEWS = 8
    CREATE_SHARED_FOLDERS_AND_VIEWS = 16
    CREATE_SHARED_FOLDERS_VIEWS = 16
    CREATE_SCRIPT_AGENTS = 32
    READ_PUBLIC_DOCUMENTS = 64
    WRITE_PUBLIC_DOCUMENTS = 128
    REPLICATE_AND_COPY_DOCUMENTS = 256
    REPLICATE_COPY_DOCUMENTS = 256


class ACLLEVEL(enum.IntEnum):
    """Access level in ACL entries"""

    NOACCESS = 0
    DEPOSITOR = 1
    READER = 2
    AUTHOR = 3
    EDITOR = 4
    DESIGNER = 5
    MANAGER = 6


class ACLTYPE(enum.IntEnum):
    """ACL Entry types"""

    UNSPECIFIED = 0
    PERSON = 1
    SERVER = 2
    MIXED_GROUP = 3
    PERSON_GROUP = 4
    SERVER_GROUP = 5


class DATECONV(enum.Flag):
    """Date conversion modes"""

    DATETIME = enum.auto()
    LOCAL = enum.auto()
    NAIVE = enum.auto()
    NATIVE = enum.auto()
    NATIVESTRING = enum.auto()
    DEFAULT = DATETIME


class RTCONV(enum.Flag):
    """Rich Text conversion types"""

    NONE = enum.auto()
    TEXT = enum.auto()
    FORMATTED = enum.auto()
    UNFORMATTED = enum.auto()
    DEFAULT = NONE
    # TODO: XML, HTML


class ITEMTYPE(enum.IntEnum):
    """Item types"""

    ACTIONCD = 16  # saved action CD records; non-Computable; canonical form.
    ASSISTANTINFO = 17  # saved assistant information; non-Computable; canonical form.
    ATTACHMENT = 1084  # file attachment.
    AUTHORS = 1076  # authors.
    COLLATION = 2  # new with Release 6.
    DATETIMES = 1024  # date-time value or range of date-time values.
    EMBEDDEDOBJECT = 1090  # embedded object.
    ERRORITEM = 256  # an error occurred while accessing the type.
    FORMULA = 1536  # Notes formula.
    HTML = 21  # HTML source text.
    ICON = 6  # icon.
    LSOBJECT = 20  # saved LotusScript Object code for an agent.
    MIME_PART = 25  # MIME support.
    NAMES = 1074  # names.
    NOTELINKS = 7  # link to a database, view, or document.
    NOTEREFS = 4  # reference to the parent document.
    NUMBERS = 768  # number or number list.
    OTHEROBJECT = 1085  # other object.
    QUERYCD = 15  # saved query CD records; non-Computable; canonical form.
    READERS = 1075  # readers.
    RFC822TEXT = 1282  # RFC822 Internet mail text.
    RICHTEXT = 1  # rich text.
    SIGNATURE = 8  # signature.
    TEXT = 1280  # text or text list.
    TEXTLIST = 1281  # text list.
    UNAVAILABLE = 512  # the item type isn't available.
    UNKNOWN = 0  # the item type isn't known.
    USERDATA = 14  # user data.
    USERID = 1792  # user ID name.
    VIEWMAPDATA = 18  # saved ViewMap dataset; non-Computable; canonical form.
    VIEWMAPLAYOUT = 19  # saved ViewMap layout; non-Computable; canonical form.
