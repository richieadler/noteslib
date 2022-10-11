"""
Document - wrapper around a NotesDocument.

"""
import datetime
import json
from functools import partial
from typing import Any, Union

import pendulum as p
from win32com.client import CDispatch

from .core import NotesLibObject, Session
from .enums import DATECONV, ITEMTYPE, RTCONV


def _c(ndt):
    return datetime.datetime.combine(ndt.date(), ndt.time(), ndt.tzinfo)


def _c_naive(ndt):
    return _c(p.instance(ndt).in_tz("local").replace(tzinfo=None))


def _c_local(ndt):
    return _c(p.instance(ndt).in_tz("local"))


def _fnstr(ndt, *, zone):
    if ndt.year < 1900:
        return None
    return p.instance(ndt).in_tz(zone).to_iso8601_string()


def _fndt(ndt, *, zone):
    if ndt.year < 1900:
        return None
    return _c(p.instance(ndt).in_tz(zone))


class Document(NotesLibObject):
    """
    The Document class encapsulates a Notes database Document. It supports all the
    properties and methods of the LotusScript NotesDocument class, using the same
    syntax.

    Additional features:

    * It can retrieve an item value using "item" syntax.

      * Using this syntax performs
        conversions on NOTESDATE fields and returns a normal ``datetime.datetime`` object
        (**not** a ``pywintype.datetime``), with GMT timezone.

       * A non-existing item raises ``KeyError`` instead of returning an array containing
         a single empty string like ``NotesDoc.GetItemValue()`` does.

      >>> doc["Form"]
      ['Memo']

    * It has a ``get()`` method which allows to get a default value for missing items,
      as well as parameters to customize the conversion of Notes DATETIME and RICHTEXT
      items.

    * The ``dict()`` method returns a ``dict`` version of the document.

    """

    def __eq__(self, other):
        docself = self.notesobj
        docother = other.notesobj
        return (
            docself.ParentDatabase.ReplicaID == docother.ParentDatabase.ReplicaID
            and docself.UniversalID == docother.UniversalID
        )

    def __getitem__(self, name):
        """Get item value by name, as index"""
        # TODO: Allow a "Compact" mode (return a scalar if there is only 1 value in the item.Values)
        # TODO: Rich text conversions
        item = self._handle.GetFirstItem(name)
        if item is None:
            raise KeyError(repr(name))
        return self.get(item)

    def get(
        self,
        item: Union[CDispatch, str],
        default: Any = None,
        *,
        convert_date: Union[DATECONV, str] = DATECONV.DEFAULT,
        convert_rt=RTCONV.NONE,
    ):
        r"""Return a list containing all values for the document item,
        with the optional conversions indicated for the relevant types.

        :param item: Document item
        :type item: NotesItem

        :param default: Value returned for a non-existing item

        :param convert_date: valid value in the enum DATECONV, or string,
            affecting NotesDateTime fields returned as DATETIME variants in COM:

            * DATECONV.DEFAULT returns Python datetime(s) corresponding to the date in UTC

            * DATECONV.LOCAL returns Python datetime(s) converted to the local time

            * DATECONV.NAIVE returns Python datetime(s) with the same date/time information
              as the original DATETIME variant converted to local time, but without
              zone information

            * DATECONV.NATIVE returns a list of ``NotesDateTime`` or ``NotesDateTimeRange`` values

            * DATECONV.NATIVESTRING returns a list of the date(s) as returned by the @Text function
              in the Notes formula language

            * String "tz:\ *zonename*\ [:str]": Convert each NotesDateTime to the zone *zonename*
              and, if `:str` is appended, convert it to the ISO-8601 representation of the date,
              otherwise return a datetime instance in the indicated zone

            Notice that if the field contains one or more `NotesDateTimeRange` objects,
            the ``.Values`` attribute for the ``NotesItem`` returns a tuple with an even number of
            DATETIME variants retrieved as ``pywintypes.datetime`` instances converted to
            localtime but with a UTC time zone; instead, ``get()`` returns a list of correctly paired
            lists containing two ``datetime.datetime`` values each, corresponding to the starting and
            ending ``datetime`` instances for each range, all with the proper time zone according to
            the specified `convert_date` value.

        :param convert_rt: conversion for ``NotesRichText`` items

        :returns: ``list`` of elements of type corresponding to the item, and the subsequent conversions

        """
        doc = self._handle
        if isinstance(item, str):
            item = doc.GetFirstItem(item)

        if not isinstance(convert_date, DATECONV) and not str(convert_date).startswith(
            "tz:"
        ):
            raise ValueError(
                'Incorrect value for parameter "convert_date": {convert_date!r}'
            )

        if item is None:
            return default
        if item.type == ITEMTYPE.RICHTEXT:
            if convert_rt == RTCONV.NONE:
                lst = [item.Text]
            elif convert_rt == RTCONV.FORMATTED:
                lst = [item.GetFormattedText(False, 120)]
            elif convert_rt == RTCONV.UNFORMATTED:
                lst = [item.GetUnformattedText()]
            else:
                raise ValueError(f"convert_rt value {convert_rt!r} is not valid")
        elif item.type == ITEMTYPE.NUMBERS:
            lst = [(int(_) if _.is_integer() else _) for _ in item.Values]
        elif item.type == ITEMTYPE.DATETIMES:
            try:
                lst = self._convert_datetime(item, convert_date)
            except OSError as err:
                db = doc.Parent
                unid = doc.UniversalID
                raise OSError(
                    f"{db.Server}!!{db.FilePath}: {unid} "
                    'Field "{item.Name}": Problem with values: {item.Values}'
                ) from err
        else:
            lst = item.Values

        return list(lst)

    def _convert_datetime(self, item, convert_date):
        doc = self._handle

        if convert_date == DATECONV.NATIVE:
            lst = item.GetValueDateTimeArray()
        elif convert_date == DATECONV.NATIVESTRING:
            ns = Session()
            lst = ns.Evaluate(f'@Text({item.Name}; "D0T0Z2")', doc)
        else:
            if convert_date == DATECONV.DEFAULT:
                func = _c
            elif convert_date == DATECONV.LOCAL:
                func = _c_local
            elif convert_date == DATECONV.NAIVE:
                func = _c_naive
            elif (sconv := str(convert_date)).startswith("tz:"):
                _, zone, conv = (sconv + ":").split(":")[:3]
                func = (
                    partial(_fnstr, zone=zone)
                    if conv == "str"
                    else partial(_fndt, zone=zone)
                )
            else:
                raise ValueError(
                    f"Value {convert_date!r} for parameter convert_date is not valid"
                )
            # Convert all dates (including ranges returned as lists)
            lst = item.GetValueDateTimeArray()
            if "DateRange" in repr(lst[0]):
                lst = [
                    [
                        func(_.StartDateTime.LSGMTTime),
                        func(_.EndDateTime.LSGMTTime),
                    ]
                    for _ in lst
                ]
            else:
                lst = [func(_.LSGMTTime) for _ in lst]
        return lst

    def dict(self, *, omit_special=False):
        """Return a ``dict`` with a reasonable representation of the document's contents.
        The dates are returned in ISO8601-compatible format.
        """
        return {
            item.Name: self.get(item, convert_date="tz:GMT:str")
            for item in self._handle.Items
            if not (omit_special and item.Name.startswith("$"))
        }

    def json(self, *, omit_special=False, **kwargs: Any):
        """Return a JSON version of the ``dict()`` method"""
        return json.dumps(self.dict(omit_special=omit_special), **kwargs)


class DocumentCollection(NotesLibObject):
    """The DocumentCollection class encapsulates a Notes database DocumentCollection. It supports
    all the properties and methods of the LotusScript NotesDocumentCollection class, using the same
    syntax.

    Additional properties:

    * The nth document of a collection can be retrieved by using ``coll[n-1]`` (indices in Python
      start at 0, NotesDocumentCollection indices start at 1). Negative indices can be used.
    * ``len(coll)`` returns the number of documents of the collection.
    * The collection can be iterated (even in reverse) so code as:

      ..  code:: python

          for doc in coll:
              print(doc["Form"])
          for doc in reversed(coll):
              print(doc.UniversalID)

      are valid.

    """

    def __init__(self, *, obj):
        super().__init__(obj=obj)

    def __len__(self):
        return self._handle.Count

    def __getitem__(self, item):
        if not isinstance(item, int):
            raise IndexError(f"{self.__class__.__name__} indices must be integers")
        ix = (item % len(self)) + 1
        doc = self._handle.GetNthDocument(ix)
        if doc.IsDeleted or not doc.IsValid:
            doc = None
        return Document(obj=doc) if doc is not None else None

    def __iter__(self):
        coll = self._handle
        doc = coll.GetFirstDocument()
        while doc:
            yield Document(obj=doc)
            doc = coll.GetNextDocument(doc)

    def __reversed__(self):
        coll = self._handle
        doc = coll.GetLastDocument()
        while doc:
            yield Document(obj=doc)
            doc = coll.GetPrevDocument(doc)
