from ctypes import POINTER
from comtypes import (
    CoInitialize,
    CoCreateInstance,
    CLSCTX_LOCAL_SERVER,
)
from .com_defns import (
    CLSID_VirtualDesktopPinnedApps,
    IVirtualDesktopManager,
    IVirtualDesktopManagerInternal,
    IVirtualDesktopManagerInternal2,
    IVirtualDesktopPinnedApps,
    IApplicationViewCollection,
    IServiceProvider,
    CLSID_ImmersiveShell,
    CLSID_VirtualDesktopManagerInternal,
)


def _get_object(cls, clsid = None):
    CoInitialize()
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

def get_vd_manager_internal2():
    return _get_object(IVirtualDesktopManagerInternal2, CLSID_VirtualDesktopManagerInternal)

def get_view_collection():
    return _get_object(IApplicationViewCollection)

def get_pinned_apps():
    return _get_object(IVirtualDesktopPinnedApps, CLSID_VirtualDesktopPinnedApps)
