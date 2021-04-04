
"""
References:
    * https://github.com/Ciantic/VirtualDesktopAccessor/blob/master/VirtualDesktopAccessor/dllmain.h
"""
from typing import List
from ctypes import windll
from comtypes.GUID import GUID

from .win10desktops import IVirtualDesktop
from .utils import get_vd_manager_internal

ASFW_ANY = -1

class VirtualDesktop():
    def __init__(
        self,
        number: int = None,
        desktop_id: GUID = None,
        desktop: IVirtualDesktop = None,
        current: bool = False
    ):
        self._manager_internal = get_vd_manager_internal()

        if number:
            if number <= 0:
                raise ValueError(f"Desktop number must be at least 1, {number} provided")
            array = self._manager_internal.GetDesktops()
            desktop_count = array.GetCount()
            if number > desktop_count:
                raise ValueError(
                    f"Desktop number {number} exceeds the number of desktops, {desktop_count}."
                )
            self._virtual_desktop = array.get_at(number - 1, IVirtualDesktop)

        elif desktop_id:
            array = self._manager_internal.GetDesktops()
            for vd in array.iter(IVirtualDesktop):
                if vd.GetID() == desktop_id:
                    self._virtual_desktop = vd
                    break
            else:
                raise Exception(f"Desktop with ID {self.id} not found")

        elif desktop:
            self._virtual_desktop = desktop

        elif current:
            self._virtual_desktop = self._manager_internal.GetCurrentDesktop()

        else:
            raise Exception("Must provide one of 'number', 'desktop_id' or 'desktop'")

    @classmethod
    def current(cls):
        """
        # TODO:
        """
        return cls(current=True)


    @property
    def id(self) -> GUID:
        """
        # TODO:
        """
        return self._virtual_desktop.GetID()

    @property
    def number(self) -> int:
        """
        # TODO:
        """
        array = self._manager_internal.GetDesktops()
        for i, vd in enumerate(array.iter(IVirtualDesktop), 1):
            if self.id == vd.GetID():
                return i
        else:
            raise Exception(f"Desktop with ID {self.id} not found")

    def go(self, allow_set_foreground: bool = True):
        """Switch to this virtual desktop.

        Args:
            allow_set_foreground (bool, optional):
                Call AllowSetForegroundWindow(ASFW_ANY) before switching.
                This partially fixes an issue where the focus remains behind after switching.
                More details [here](https://github.com/Ciantic/VirtualDesktopAccessor/issues/4) and [here](https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-allowsetforegroundwindow).
                Defaults to True.
        """
        if allow_set_foreground:
            windll.user32.AllowSetForegroundWindow(ASFW_ANY)
        self._manager_internal.SwitchDesktop(self._virtual_desktop)


def get_virtual_desktops() -> List[VirtualDesktop]:
    """
    # TODO:
    """
    manager_internal = get_vd_manager_internal()
    array = manager_internal.GetDesktops()
    return [VirtualDesktop(desktop=vd) for vd in array.iter(IVirtualDesktop)]

