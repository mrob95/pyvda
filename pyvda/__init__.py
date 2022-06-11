"""
Python Virtual Desktop Accessor
================================

Implements a subset of the functionality offered by https://github.com/Ciantic/VirtualDesktopAccessor, written in pure Python and installable via pip.

Installation
-------------
.. code:: shell

    $ pip install pyvda

Example
-------
.. code:: python

    from pyvda import AppView, get_apps_by_z_order, VirtualDesktop, get_virtual_desktops

    number_of_active_desktops = len(get_virtual_desktops())
    print(f"There are {number_of_active_desktops} active desktops")

    current_desktop = VirtualDesktop.current()
    print(f"Current desktop is number {current_desktop}")

    current_window = AppView.current()
    target_desktop = VirtualDesktop(5)
    current_window.move(target_desktop)
    print(f"Moved window {current_window.hwnd} to {target_desktop.number}")

    print("Going to desktop number 5")
    VirtualDesktop(5).go()

    print("Pinning the current window")
    AppView.current().pin()
"""

__version__ = "0.2.7"

import platform
import os

def _check_version():
    if platform.system() != "Windows" or platform.release() != "10":
        raise WindowsError(
            "The virtual desktop feature is only available on Windows 10"
        )

if not os.getenv("READTHEDOCS"):
    _check_version()

from .pyvda import (
    AppView,
    get_apps_by_z_order,
    VirtualDesktop,
    get_virtual_desktops,
    set_wallpaper_for_all_desktops,
)
