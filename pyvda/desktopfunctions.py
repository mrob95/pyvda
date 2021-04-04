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

