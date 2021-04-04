from ctypes import POINTER, byref

from datetime import datetime
from ctypes import POINTER, byref
from ctypes.wintypes import BOOL, LPCWSTR, LPWSTR
import platform
from typing import Iterator
from comtypes import (
    CoCreateInstance,
    CLSCTX_LOCAL_SERVER,
    IUnknown,
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

def get_vd_manager():
    return _get_object(IVirtualDesktopManager)

def get_vd_manager_internal():
    return _get_object(IVirtualDesktopManagerInternal, CLSID_VirtualDesktopManagerInternal)

def get_view_collection():
    return _get_object(IApplicationViewCollection)

def get_pinned_apps():
    return _get_object(IVirtualDesktopPinnedApps, CLSID_VirtualDesktopPinnedApps)
