import platform

def _check_version():
    if platform.system() != "Windows" or platform.release() != "10":
        raise WindowsError(
            "The virtual desktop feature is only available on Windows 10"
        )

_check_version()

from .app_view import AppView, get_apps_by_z_order
from .desktop import VirtualDesktop, get_virtual_desktops