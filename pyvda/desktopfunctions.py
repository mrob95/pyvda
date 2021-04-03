"""
References:
    * https://github.com/Ciantic/VirtualDesktopAccessor/blob/master/VirtualDesktopAccessor/dllmain.h
"""
from datetime import datetime
from ctypes import POINTER, byref
from ctypes.wintypes import BOOL, LPCWSTR, LPWSTR
import platform
from comtypes import (
    CoCreateInstance,
    CLSCTX_LOCAL_SERVER,
)
from .win10desktops import (
    CLSID_VirtualDesktopPinnedApps, IVirtualDesktop,
    IVirtualDesktopManager,
    IVirtualDesktopManagerInternal,
    IApplicationView,
    IVirtualDesktopPinnedApps,
    IObjectArray,
    IApplicationViewCollection,
    IServiceProvider,
    CLSID_ImmersiveShell,
    CLSID_IVirtualDesktopManager,
    CLSID_VirtualDesktopManagerInternal,
)

def _get_object(cls, clsid = None):
    pServiceProvider = CoCreateInstance(
        CLSID_ImmersiveShell, IServiceProvider, CLSCTX_LOCAL_SERVER
    )
    pObject = POINTER(cls)()
    pServiceProvider.QueryService(
        clsid or cls._iid_,
        cls._iid_,
        pObject,
    )
    return pObject

def _get_vd_manager():
    return _get_object(IVirtualDesktopManager)

def _get_vd_manager_internal():
    return _get_object(IVirtualDesktopManagerInternal, CLSID_VirtualDesktopManagerInternal)

def _get_view_collection():
    return _get_object(IApplicationViewCollection)

def _get_pinned_apps():
    return _get_object(IVirtualDesktopPinnedApps, CLSID_VirtualDesktopPinnedApps)


def _get_application_view_for_hwnd(hwnd):
    # Get the IApplicationView for the window
    pViewCollection = _get_view_collection()
    pView = pViewCollection.GetViewForHwnd(hwnd)
    return pView

def _get_application_id_for_hwnd(hwnd):
    pView = _get_application_view_for_hwnd(hwnd)
    app_id = pView.GetAppUserModelId()
    return app_id


def _get_desktop_by_id(id):
    pManagerInternal = _get_vd_manager_internal()
    array = pManagerInternal.GetDesktops()
    for i in range(array.GetCount()):
        item = POINTER(IVirtualDesktop)()
        array.GetAt(i, IVirtualDesktop._iid_, item)
        item_id = item.GetID()
        if item_id == id:
            return item


def _get_desktop_number(desktop):
    pManagerInternal = _get_vd_manager_internal()
    target_id = desktop.GetID()
    array = pManagerInternal.GetDesktops()
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
    array = pManagerInternal.GetDesktops()
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
    pApplicationView = _get_application_view_for_hwnd(hwnd)

    # Get the IVirtualDesktop for the target desktop
    array = pManagerInternal.GetDesktops()
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
    array = pManagerInternal.GetDesktops()

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
    pApplicationView = _get_application_view_for_hwnd(hwnd)
    desktopId = pApplicationView.GetVirtualDesktopId()
    desktop = _get_desktop_by_id(desktopId)
    desktop_number = _get_desktop_number(desktop) + 1
    return desktop_number


# def PinWindow(hwnd: int) -> None:
def PinWindow(hwnd):
    """
    Pin a window (corresponds to the 'show window on all desktops' toggle).

    Arguments:
        hwnd {int} -- Handle of the window to pin, from e.g. win32gui.GetForegroundWindow().
    """
    pinnedApps = _get_pinned_apps()
    pView = _get_application_view_for_hwnd(hwnd)
    pinnedApps.PinView(pView)


# def UnPinWindow(hwnd: int) -> None:
def UnPinWindow(hwnd):
    """
    Unpin a window (corresponds to the 'show window on all desktops' toggle).

    Arguments:
        hwnd {int} -- Handle of the window to unpin, from e.g. win32gui.GetForegroundWindow().
    """
    pinnedApps = _get_pinned_apps()
    pView = _get_application_view_for_hwnd(hwnd)
    pinnedApps.UnpinView(pView)


# def IsPinnedWindow(hwnd: int) -> bool:
def IsPinnedWindow(hwnd):
    """
    Check if a window is pinned (corresponds to the 'show window on all desktops' toggle).

    Arguments:
        hwnd {int} -- Handle of the window to check, from e.g. win32gui.GetForegroundWindow().

    Returns:
        bool -- is the window pinned?.
    """
    pinnedApps = _get_pinned_apps()
    pView = _get_application_view_for_hwnd(hwnd)
    isPinned = pinnedApps.IsViewPinned(pView)
    return isPinned


# def PinApp(hwnd: int) -> None:
def PinApp(hwnd):
    """
    Pin an app (corresponds to the 'show windows from this app on all desktops' toggle).

    Arguments:
        hwnd {int} -- Handle of a window belonging to the app to pin, from e.g. win32gui.GetForegroundWindow().
    """
    pinnedApps = _get_pinned_apps()
    app_id = _get_application_id_for_hwnd(hwnd)
    pinnedApps.PinAppID(app_id)


# def UnPinApp(hwnd: int) -> None:
def UnPinApp(hwnd):
    """
    Unpin an app (corresponds to the 'show windows from this app on all desktops' toggle).

    Arguments:
        hwnd {int} -- Handle of a window belonging to the app to unpin, from e.g. win32gui.GetForegroundWindow().
    """
    pinnedApps = _get_pinned_apps()
    app_id = _get_application_id_for_hwnd(hwnd)
    pinnedApps.UnpinAppID(app_id)


# def IsPinnedApp(hwnd: int) -> bool:
def IsPinnedApp(hwnd):
    """
    Check if an app is pinned (corresponds to the 'show windows from this app on all desktops' toggle).

    Arguments:
        hwnd {int} -- Handle of a window belonging to the app to check, from e.g. win32gui.GetForegroundWindow().

    Returns:
        bool -- is the app pinned?.
    """
    pinnedApps = _get_pinned_apps()
    app_id = _get_application_id_for_hwnd(hwnd)
    isPinned = pinnedApps.IsAppIdPinned(app_id)
    return isPinned


# def ViewIsShownInSwitchers(hwnd: int) -> bool:
def ViewIsShownInSwitchers(hwnd):
    """
    Arguments:
        hwnd {int}

    Returns:
        bool -- is the app shown in the alt-tab view?
    """
    view = _get_application_view_for_hwnd(hwnd)
    return bool(view.GetShowInSwitchers())


# def ViewIsVisible(hwnd: int) -> bool:
def ViewIsVisible(hwnd):
    """
    Arguments:
        hwnd {int}

    Returns:
        bool -- is the app visible?
    """
    view = _get_application_view_for_hwnd(hwnd)
    return bool(view.GetVisibility())


# def ViewGetLastActivationTimestamp(hwnd: int) -> int:
def ViewGetLastActivationTimestamp(hwnd):
    """
    Arguments:
        hwnd {int}

    Returns:
        int -- last activation timestamp
    """
    view = _get_application_view_for_hwnd(hwnd)
    return view.GetLastActivationTimestamp()

# def ViewSetFocus(hwnd: int) -> int:
def ViewSetFocus(hwnd):
    """Focus a window

    Arguments:
        hwnd {int} -- Handle to the window to focus
    """
    view = _get_application_view_for_hwnd(hwnd)
    return view.SetFocus()

# def ViewSwitchTo(hwnd: int) -> int:
def ViewSwitchTo(hwnd):
    """Switch to a window. Behaves slightly differently to ViewSetFocus -
    this is what is called when you use the alt-tab menu.

    Arguments:
        hwnd {int} -- Handle to the window to focus
    """
    view = _get_application_view_for_hwnd(hwnd)
    return view.SwitchTo()


# def ViewGetByZOrder(switcher_windows: bool = True, current_desktop: bool = True) -> List[int]:
def ViewGetByZOrder(switcher_windows = True, current_desktop = True):
    """Get a list of window handles, ordered by their Z position, with
    the foreground window first.

    Arguments:
        switcher_windows {bool} -- Only include windows which appear in the alt-tab dialogue
        current_desktop {bool} -- Only include windows which are on the current virtual desktop

    Returns:
        List[int] -- Window handles
    """
    collection = _get_view_collection()
    vdm = _get_vd_manager()
    views_arr = collection.GetViewsByZOrder()
    result = []
    for i in range(views_arr.GetCount()):
        item = POINTER(IApplicationView)()
        views_arr.GetAt(i, IApplicationView._iid_, item)
        view = item.GetThumbnailWindow()

        if switcher_windows and not item.GetShowInSwitchers():
            continue
        if current_desktop and not vdm.IsWindowOnCurrentVirtualDesktop(view):
            continue

        result.append(view)

    return result

