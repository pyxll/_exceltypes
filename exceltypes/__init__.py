"""Excel Types win32com Extension.

This package provides pythoncom types for Excel COM classes.

For most cases using the win32com makepy types with Excel works well,
but since Excel 2007 the callback objects implementing the IRTDUpdateEvent
interface passed to IRtdServer RTD server implementations doesn't work
via the usual IDispatch mechanism.

This package registers extension types with pythoncom so they can be called
directly rather than relying on the win32com makepy classes which make all calls
via the IDispatch interface.

Example::

    import exceltypes

    # cast an object to a PyIRTDUpdateEvent instance using QueryInterface
    obj = <a PyIDispatch object passed from Excel>
    update_event = obj.QueryInterface(exceltypes.IID_IRTDUpdateEvent)

References:
    - http://stackoverflow.com/questions/9985512/excel-rtd-server-in-python-not-updating-data
    - https://mail.python.org/pipermail/python-win32/2012-April/012207.html
"""
import pythoncom
from ._exceltypes import *

__all__ = [
    "IID_IRTDUpdateEvent",
]
