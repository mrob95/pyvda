from ctypes import POINTER
import platform
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


def _check_version():
    if platform.system() != "Windows" or platform.release() != "10":
        raise WindowsError(
            "The virtual desktop feature is only available on Windows 10"
        )


# def GetCurrentDesktopNumber() -> int:
def GetCurrentDesktopNumber():
    """
    Returns:
        int -- The number of the virtual desktop you're currently on, with the first being number 1 etc.
    """
    _check_version()
    pManagerInternal = _get_vd_manager_internal()
    currentDesktop = POINTER(IVirtualDesktop)()
    pManagerInternal.GetCurrentDesktop(currentDesktop)
    number = _get_desktop_number(currentDesktop)
    # Account for zero indexing
    return number + 1


# def GetDesktopCount() -> int:
def GetDesktopCount():
    """
    Returns:
        int -- The number of virtual desktops active.
    """
    _check_version()
    pManagerInternal = _get_vd_manager_internal()
    array = POINTER(IObjectArray)()
    pManagerInternal.GetDesktops(array)
    count = array.GetCount()
    return count


# def MoveWindowToDesktopNumber(hwnd: int, number: int) -> None:
def MoveWindowToDesktopNumber(hwnd, number):
    """
    Move a window to a different virtual desktop.

    Arguments:
        hwnd {int} -- Handle of the window to be moved, from e.g. win32gui.GetForegroundWindow().
        number {int} -- Desktop number to move the window to, between 1 and the number of desktops active.
    """
    _check_version()
    pManagerInternal = _get_vd_manager_internal()
    pViewCollection = _get_view_collection()
    # Get the IApplicationView for the window
    pApplicationView = POINTER(IApplicationView)()
    pViewCollection.GetViewForHwnd(hwnd, pApplicationView)

    # Get the IVirtualDesktop for the target desktop
    array = POINTER(IObjectArray)()
    pManagerInternal.GetDesktops(array)
    if number <= 0:
        raise ValueError("Desktop number must be at least 1, %s provided" % number)
    desktop_count = array.GetCount()
    if number > desktop_count:
        raise ValueError(
            "Desktop number %s exceeds the number of desktops, %s." % (number, desktop_count),
        )
    desktop = POINTER(IVirtualDesktop)()
    # -1 to get correct index
    array.GetAt(number - 1, IVirtualDesktop._iid_, desktop)

    pManagerInternal.MoveViewToDesktop(pApplicationView, desktop)


# def GoToDesktopNumber(number: int) -> None:
def GoToDesktopNumber(number):
    """
    Switch to a different virtual desktop.

    Arguments:
        number {int} -- Desktop number to switch to.
    """
    _check_version()
    pManagerInternal = _get_vd_manager_internal()
    array = POINTER(IObjectArray)()
    pManagerInternal.GetDesktops(array)

    if number <= 0:
        raise ValueError("Desktop number must be at least 1, %s provided" % number)
    desktop_count = array.GetCount()
    if number > desktop_count:
        raise ValueError(
            "Desktop number %s exceeds the number of desktops, %s." % (number, desktop_count),
        )

    desktop = POINTER(IVirtualDesktop)()
    array.GetAt(number - 1, IVirtualDesktop._iid_, desktop)
    pManagerInternal.SwitchDesktop(desktop)


# def GetWindowDesktopNumber(hwnd: int) -> int:
def GetWindowDesktopNumber(hwnd):
    """
    Returns the number of the desktop which a particular window is on.

    Arguments:
        hwnd {int} -- Handle of the window, from e.g. win32gui.GetForegroundWindow().

    Returns:
        int -- Its desktop number.
    """
    _check_version()
    pViewCollection = _get_view_collection()
    app = POINTER(IApplicationView)()
    pViewCollection.GetViewForHwnd(hwnd, app)
    desktopId = app.GetVirtualDesktopId()
    desktop = _get_desktop_by_id(desktopId)
    desktop_number = _get_desktop_number(desktop) + 1
    return desktop_number
