# comtypes spams debug output for every allocation
import logging
logging.getLogger("comtypes").setLevel(logging.WARNING)

from .desktopfunctions import (
    GetCurrentDesktopNumber,
    GetDesktopCount,
    MoveWindowToDesktopNumber,
    GoToDesktopNumber,
    GetWindowDesktopNumber,
)
