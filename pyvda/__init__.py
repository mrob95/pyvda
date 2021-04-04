import platform


def _check_version():
    if platform.system() != "Windows" or platform.release() != "10":
        raise WindowsError(
            "The virtual desktop feature is only available on Windows 10"
        )

_check_version()


# from .desktopfunctions import (
#     GetCurrentDesktopNumber,
#     GetDesktopCount,
#     MoveWindowToDesktopNumber,
#     GoToDesktopNumber,
#     GetWindowDesktopNumber,
#     PinWindow,
#     UnPinWindow,
#     IsPinnedWindow,
#     PinApp,
#     UnPinApp,
#     IsPinnedApp,
#     ViewIsShownInSwitchers,
#     ViewIsVisible,
#     ViewGetByZOrder,
#     ViewGetLastActivationTimestamp,
#     ViewSetFocus,
#     ViewSwitchTo,
# )
from .app_view import AppView
from .desktop import VirtualDesktop, get_virtual_desktops