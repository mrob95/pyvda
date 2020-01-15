from ctypes import POINTER
import sys
from comtypes import (
    IUnknown,
    GUID,
    IID,
    COMMETHOD,
    CoInitialize,
    CoCreateInstance,
    cast,
    CLSCTX_LOCAL_SERVER,
)
from .win10desktops import (
    IVirtualDesktop,
    IVirtualDesktopManager,
    IVirtualDesktopManagerInternal,
    IApplicationView,
    IObjectArray,
    IApplicationViewCollection,
    IServiceProvider,
    CLSID_ImmersiveShell,
    CLSID_IVirtualDesktopManager,
    CLSID_VirtualDesktopManagerInternal,
)


def _get_vd_manager_internal():
    CoInitialize()
    pServiceProvider = CoCreateInstance(
        CLSID_ImmersiveShell, IServiceProvider, CLSCTX_LOCAL_SERVER
    )
    pManagerInternal = POINTER(IVirtualDesktopManagerInternal)()
    pServiceProvider.QueryService(
        CLSID_VirtualDesktopManagerInternal,
        IVirtualDesktopManagerInternal._iid_,
        pManagerInternal,
    )
    return pManagerInternal


def _get_view_collection():
    pServiceProvider = CoCreateInstance(
        CLSID_ImmersiveShell, IServiceProvider, CLSCTX_LOCAL_SERVER
    )
    pViewCollection = POINTER(IApplicationViewCollection)()
    pServiceProvider.QueryService(
        IApplicationViewCollection._iid_,
        IApplicationViewCollection._iid_,
        pViewCollection,
    )
    return pViewCollection


def _get_desktop_by_number(number):
    pManagerInternal = _get_vd_manager_internal()
    array = POINTER(IObjectArray)()
    pManagerInternal.GetDesktops(array)
    if number < 0:
        raise ValueError(
            f"Desktop number must be at least zero, {number} provided"
        )
    desktop_count = array.GetCount()
    if number >= desktop_count:
        raise ValueError(
            f"Desktop number {number} exceeds the maximum allowed, {desktop_count - 1} (remember that desktops are zero indexed)",
        )
    found = POINTER(IVirtualDesktop)()
    array.GetAt(number, IVirtualDesktop._iid_, found)
    return found


def _get_desktop_by_id(id):
    pManagerInternal = _get_vd_manager_internal()
    array = POINTER(IObjectArray)()
    pManagerInternal.GetDesktops(array)
    for i in range(array.GetCount()):
        item = POINTER(IVirtualDesktop)()
        array.GetAt(i, IVirtualDesktop._iid_, item)
        item_id = item.GetID()
        if item_id == id:
            return item


def _check_version():
    if sys.getwindowsversion().major != 10:
        raise WindowsError("The virtual desktop feature is only available on Windows 10")


def _get_desktop_number(desktop):
    pManagerInternal = _get_vd_manager_internal()
    target_id = desktop.GetID()
    array = POINTER(IObjectArray)()
    pManagerInternal.GetDesktops(array)
    for i in range(array.GetCount()):
        item = POINTER(IVirtualDesktop)()
        array.GetAt(i, IVirtualDesktop._iid_, item)
        item_id = item.GetID()
        if item_id == target_id:
            return i


def GetCurrentDesktopNumber():
    _check_version()
    pManagerInternal = _get_vd_manager_internal()
    currentDesktop = POINTER(IVirtualDesktop)()
    pManagerInternal.GetCurrentDesktop(currentDesktop)
    number = _get_desktop_number(currentDesktop)
    return number


def GetDesktopCount():
    _check_version()
    pManagerInternal = _get_vd_manager_internal()
    array = POINTER(IObjectArray)()
    pManagerInternal.GetDesktops(array)
    count = array.GetCount()
    return count


def MoveWindowToDesktopNumber(hwnd, number):
    _check_version()
    pManagerInternal = _get_vd_manager_internal()
    pViewCollection = _get_view_collection()
    pApplicationView = POINTER(IApplicationView)()
    pViewCollection.GetViewForHwnd(hwnd, pApplicationView)
    desktop = _get_desktop_by_number(number)
    pManagerInternal.MoveViewToDesktop(pApplicationView, desktop)


def GoToDesktopNumber(number):
    _check_version()
    pManagerInternal = _get_vd_manager_internal()
    array = POINTER(IObjectArray)()
    pManagerInternal.GetDesktops(array)

    if number < 0:
        raise ValueError(
            f"Desktop number must be at least zero, {number} provided"
        )
    desktop_count = array.GetCount()
    if number >= desktop_count:
        raise ValueError(
            f"Desktop number {number} exceeds the maximum allowed, {desktop_count - 1} (remember that desktops are zero indexed)",
        )

    desktop = POINTER(IVirtualDesktop)()
    array.GetAt(number, IVirtualDesktop._iid_, desktop)
    pManagerInternal.SwitchDesktop(desktop)


def GetWindowDesktopNumber(hwnd):
    _check_version()
    pViewCollection = _get_view_collection()
    app = POINTER(IApplicationView)()
    pViewCollection.GetViewForHwnd(hwnd, app)
    desktopId = app.GetVirtualDesktopId()
    desktop = _get_desktop_by_id(desktopId)
    desktop_number = _get_desktop_number(desktop)
    return desktop_number
