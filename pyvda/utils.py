import logging
import threading
from ctypes import POINTER
from comtypes import (
    CoInitializeEx,
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
    BUILD_OVER_21313,
)

logger = logging.getLogger(__name__)


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

def get_vd_manager_internal2():
    return _get_object(IVirtualDesktopManagerInternal2, CLSID_VirtualDesktopManagerInternal)

def get_view_collection():
    return _get_object(IApplicationViewCollection)

def get_pinned_apps():
    return _get_object(IVirtualDesktopPinnedApps, CLSID_VirtualDesktopPinnedApps)

class Managers(threading.local):
    def __init__(self):
        self.try_init_com()
        self.manager_internal = get_vd_manager_internal()
        self.view_collection = get_view_collection()
        self.pinned_apps = get_pinned_apps()

        # Old interface only used for SetName
        if not BUILD_OVER_21313:
            self.manager_internal2 = get_vd_manager_internal2()

    @staticmethod
    def try_init_com():
        try:
            CoInitializeEx()
        except OSError as e:
            # This is likely because COM has already been initialised with
            # COINIT_MULTITHREADED, whereas CoInitialize uses COINIT_APARTMENTTHREADED.
            # This should not be a problem for us, so warn and keep going.
            logger.warning("Failed to initialise COM: %s", str(e))
