import logging
import sys
import threading
from ctypes import POINTER
from typing import Optional

import _ctypes
from comtypes import CLSCTX_LOCAL_SERVER, CoCreateInstance, CoInitializeEx

from pyvda.com_base import IServiceProvider
from pyvda.com_defns import (
    CLSID_ImmersiveShell,
    CLSID_VirtualDesktopManagerInternal,
    CLSID_VirtualDesktopPinnedApps,
    IApplicationViewCollection,
    IVirtualDesktopManagerInternal,
    IVirtualDesktopManagerInternal2,
    IVirtualDesktopPinnedApps,
)

logger = logging.getLogger(__name__)


def _get_object(cls, clsid = None):
    try:
        pServiceProvider = CoCreateInstance(
            CLSID_ImmersiveShell, IServiceProvider, CLSCTX_LOCAL_SERVER
        )
        pObject = POINTER(cls)()
        pServiceProvider.QueryService( # type: ignore
            clsid or cls._iid_,
            cls._iid_,
            pObject,
        )
    except _ctypes.COMError as e:
        winver = sys.getwindowsversion()
        platver = sys.getwindowsversion().platform_version
        raise NotImplementedError(
            f"Interface {cls.__name__} with ID {cls._iid_} not supported for windows version {winver.major}.{winver.minor}.{winver.build}, platform version {platver[0]}.{platver[1]}.{platver[2]}. Please open an issue at https://github.com/mrob95/pyvda/issues."
        )
    return pObject

def get_vd_manager_internal():
    return _get_object(IVirtualDesktopManagerInternal, CLSID_VirtualDesktopManagerInternal)

def get_vd_manager_internal2() -> Optional[IVirtualDesktopManagerInternal2]:
    try:
        return _get_object(IVirtualDesktopManagerInternal2, CLSID_VirtualDesktopManagerInternal) # type: ignore
    except NotImplementedError:
        return None

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
