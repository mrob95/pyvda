from typing import List
from ctypes import windll
from comtypes import GUID

from .win10desktops import IVirtualDesktop
from .utils import get_vd_manager_internal

ASFW_ANY = -1

class VirtualDesktop():
    """
    Wrapper around the `IVirtualDesktop` COM object, representing one virtual desktop.
    """
    def __init__(
        self,
        number: int = None,
        desktop_id: GUID = None,
        desktop: 'IVirtualDesktop' = None,
        current: bool = False
    ):
        """One of the following arguments must be provided:

        Args:
            number (int, optional): The number of the desired desktop in the task view (1-indexed). Defaults to None.
            desktop_id (GUID, optional): A desktop GUID. Defaults to None.
            desktop (IVirtualDesktop, optional): An `IVirtualDesktop`. Defaults to None.
            current (bool, optional): The current virtual desktop. Defaults to False.
        """
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
        """Convenience method to return a `VirtualDesktop` object for the
        currently active desktop.

        Returns:
            VirtualDesktop: The current desktop.
        """
        return cls(current=True)


    @property
    def id(self) -> GUID:
        """The GUID of this desktop.

        Returns:
            GUID: The unique id for this desktop.
        """
        return self._virtual_desktop.GetID()

    @property
    def number(self) -> int:
        """The index of this virtual desktop in the task view. Between 1 and
        the total number of desktops active.

        Returns:
            int: The desktop number.
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
            allow_set_foreground (bool, optional): Call AllowSetForegroundWindow(ASFW_ANY) before switching. This partially fixes an issue where the focus remains behind after switching. Defaults to True.

        Note:
            More details at https://github.com/Ciantic/VirtualDesktopAccessor/issues/4 and https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-allowsetforegroundwindow.
        """
        if allow_set_foreground:
            windll.user32.AllowSetForegroundWindow(ASFW_ANY)
        self._manager_internal.SwitchDesktop(self._virtual_desktop)

    # TODO: List windows on this desktop?
    # would be similar to get_by_z


def get_virtual_desktops() -> List[VirtualDesktop]:
    """Return a list of all current virtual desktops, one for each desktop visible in the task view.

    Returns:
        List[VirtualDesktop]: Virtual desktops currently active.
    """
    manager_internal = get_vd_manager_internal()
    array = manager_internal.GetDesktops()
    return [VirtualDesktop(desktop=vd) for vd in array.iter(IVirtualDesktop)]

