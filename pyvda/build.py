import logging
import os
import sys
from ctypes import POINTER

import _ctypes
from comtypes import CLSCTX_LOCAL_SERVER, GUID, CoCreateInstance, IUnknown

import pyvda.const as const
from pyvda.com_base import IServiceProvider

logger = logging.getLogger(__name__)


OVER_19041 = False
OVER_20231 = False
OVER_21313 = False
OVER_22449 = False
OVER_22621 = False
OVER_22631 = False

def try_create_manager(guid: GUID) -> bool:
    pServiceProvider = CoCreateInstance(
        const.CLSID_ImmersiveShell, IServiceProvider, CLSCTX_LOCAL_SERVER
    )

    pObject = POINTER(IUnknown)()
    try:
        pServiceProvider.QueryService( # type: ignore
            const.CLSID_VirtualDesktopManagerInternal,
            guid,
            pObject,
        )
    except _ctypes.COMError as e:
        logger.debug(f"Querying {guid}... {e.text}")
        return False
    logger.debug(f"Querying {guid}... Success!")
    return True

def do_feature_detection():
    global OVER_19041
    global OVER_20231
    global OVER_21313
    global OVER_22449
    global OVER_22621
    global OVER_22631

    if os.getenv("READTHEDOCS"):
        return

    logger.debug("Starting feature detection...")

    if try_create_manager(const.GUID_IVirtualDesktopManagerInternal_22631):
        logger.debug("Feature detection complete. Windows version is over 22631")
        OVER_19041 = True
        OVER_20231 = True
        OVER_21313 = True
        OVER_22449 = True
        OVER_22621 = True
        OVER_22631 = True
        return

    if try_create_manager(const.GUID_IVirtualDesktopManagerInternal_22621):
        logger.debug("Feature detection complete. Windows version is over 22621")
        OVER_19041 = True
        OVER_20231 = True
        OVER_21313 = True
        OVER_22449 = True
        OVER_22621 = True
        return

    if try_create_manager(const.GUID_IVirtualDesktopManagerInternal_21313):
        OVER_19041 = True
        OVER_20231 = True
        OVER_21313 = True
        # ideally we would avoid this, but 22449 changed
        # a method without updating the guid.
        build = sys.getwindowsversion().build
        if build >= 22449:
            logger.debug("Feature detection complete. Windows version is over 22449")
            OVER_22449 = True
            return
        logger.debug(f"Feature detection complete. Windows version is over 21313 (build was {build})")
        return

    if try_create_manager(const.GUID_IVirtualDesktopManagerInternal_20231):
        logger.debug("Feature detection complete. Windows version is over 20231")
        OVER_19041 = True
        OVER_20231 = True
        return

    if not try_create_manager(const.GUID_IVirtualDesktopManagerInternal_9000):
        raise NotImplementedError(
            f"""
    No supported IVirtualDesktopManagerInternal interface found.
        * Windows version is {sys.getwindowsversion()}
        * Platform version is {sys.getwindowsversion().platform_version}
    Please run with debug logging enabled and then open an issue at https://github.com/mrob95/pyvda/issues containing the complete output."""
        )

    if sys.getwindowsversion().build >= 19041:
        logger.debug("Feature detection complete. Windows version is over 19041")
        OVER_19041 = True
        return

    logger.debug("Feature detection complete. Windows version is under 19041")

do_feature_detection()
