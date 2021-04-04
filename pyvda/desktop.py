
"""
References:
    * https://github.com/Ciantic/VirtualDesktopAccessor/blob/master/VirtualDesktopAccessor/dllmain.h
"""
from typing import List
from comtypes.GUID import GUID
from .win10desktops import (
    IVirtualDesktop,
)
from .utils import (
    get_vd_manager_internal,
)

class VirtualDesktop():
    def __init__(self, number: int = None, desktop_id: GUID = None, desktop: IVirtualDesktop = None):
        if number:
            if number <= 0:
                raise ValueError(f"Desktop number must be at least 1, {number} provided")
            manager_internal = get_vd_manager_internal()
            array = manager_internal.GetDesktops()
            desktop_count = array.GetCount()
            if number > desktop_count:
                raise ValueError(
                    f"Desktop number {number} exceeds the number of desktops, {desktop_count}."
                )
            self._virtual_desktop = array.get_at(number - 1, IVirtualDesktop)

        elif desktop_id:
            manager_internal = get_vd_manager_internal()
            array = manager_internal.GetDesktops()
            for vd in array.iter(IVirtualDesktop):
                if vd.GetID() == desktop_id:
                    self._virtual_desktop = vd
                    break
            else:
                raise Exception(f"Desktop with ID {self.id} not found")

        elif desktop:
            self._virtual_desktop = desktop

        else:
            raise Exception("Must provide one of 'number', 'desktop_id' or 'desktop'")

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
        manager_internal = get_vd_manager_internal()
        array = manager_internal.GetDesktops()
        for i, vd in enumerate(array.iter(IVirtualDesktop), 1):
            if self.id == vd.GetID():
                return i
        else:
            raise Exception(f"Desktop with ID {self.id} not found")

    @classmethod
    def current(cls):
        """
        # TODO:
        """
        manager_internal = get_vd_manager_internal()
        current = manager_internal.GetCurrentDesktop()
        return cls(desktop=current)

    def go(self):
        """
        Switch to this virtual desktop.
        """
        manager_internal = get_vd_manager_internal()
        manager_internal.SwitchDesktop(self._virtual_desktop)


def get_virtual_desktops() -> List[VirtualDesktop]:
    """
    # TODO:
    """
    manager_internal = get_vd_manager_internal()
    array = manager_internal.GetDesktops()
    return [VirtualDesktop(desktop=vd) for vd in array.iter(IVirtualDesktop)]

