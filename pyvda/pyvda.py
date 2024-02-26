from __future__ import annotations

from ctypes import windll
from typing import List, Optional

from comtypes import GUID

import pyvda.build as build
from pyvda.com_defns import IApplicationView, IVirtualDesktop, IVirtualDesktop2
from pyvda.utils import Managers
from pyvda.winstring import HSTRING

ASFW_ANY = -1
NULL_PTR = 0

managers = Managers()


class AppView():
    """
    A wrapper around an `IApplicationView` object exposing window functionality relating to:

        * Setting focus
        * Pinning and unpinning (making a window persistent across all virtual desktops)
        * Moving a window between virtual desktops

    """

    def __init__(self, hwnd: Optional[int] = None, view: Optional['IApplicationView'] = None):
        """One of the following parameters must be provided:

        Args:
            hwnd (int, optional): Handle to a window. Defaults to None.
            view (IApplicationView, optional): An `IApplicationView` object. Defaults to None.
        """
        if hwnd:
            # Get the IApplicationView for the window
            self._view = managers.view_collection.GetViewForHwnd(hwnd) # type: ignore
        elif view:
            self._view = view
        else:
            raise Exception(f"Must pass 'hwnd' or 'view'")

    def __eq__(self, other):
        return self.hwnd == other.hwnd

    @property
    def hwnd(self) -> int:
        """This window's handle.
        """
        return self._view.GetThumbnailWindow() # type: ignore

    @property
    def app_id(self) -> int:
        """The ID of this window's app.
        """
        return self._view.GetAppUserModelId() # type: ignore

    @classmethod
    def current(cls):
        """
        Returns:
            AppView: An AppView for the currently focused window.
        """
        focused = managers.view_collection.GetViewInFocus() # type: ignore
        return cls(view=focused)

    #  ------------------------------------------------
    #  IApplicationView methods
    #  ------------------------------------------------
    def is_shown_in_switchers(self) -> bool:
        """Is the view shown in the alt-tab view?
        """
        return bool(self._view.GetShowInSwitchers()) # type: ignore

    def is_visible(self) -> bool:
        """Is the view visible?
        """
        return bool(self._view.GetVisibility()) # type: ignore

    def get_activation_timestamp(self) -> int:
        """Get the last activation timestamp for this window.
        """
        return self._view.GetLastActivationTimestamp() # type: ignore

    def set_focus(self, allow_set_foreground: bool = True):
        """Focus the window

        Args:
            allow_set_foreground (bool, optional): Call AllowSetForegroundWindow(ASFW_ANY) before switching. This partially fixes an issue where the focus remains behind after switching. Defaults to True.
        """
        if allow_set_foreground:
            windll.user32.AllowSetForegroundWindow(ASFW_ANY)
        return self._view.SetFocus() # type: ignore

    def switch_to(self):
        """Switch to the window. Behaves slightly differently to set_focus -
        this is what is called when you use the alt-tab menu."""
        return self._view.SwitchTo() # type: ignore


    #  ------------------------------------------------
    #  IVirtualDesktopPinnedApps methods
    #  ------------------------------------------------
    def pin(self):
        """
        Pin the window (corresponds to the 'show window on all desktops' toggle).
        """
        managers.pinned_apps.PinView(self._view) # type: ignore

    def unpin(self):
        """
        Unpin the window (corresponds to the 'show window on all desktops' toggle).
        """
        managers.pinned_apps.UnpinView(self._view) # type: ignore

    def is_pinned(self) -> bool:
        """
        Check if this window is pinned (corresponds to the 'show window on all desktops' toggle).

        Returns:
            bool: is the window pinned?
        """
        return managers.pinned_apps.IsViewPinned(self._view) # type: ignore

    def pin_app(self):
        """
        Pin this window's app (corresponds to the 'show windows from this app on all desktops' toggle).
        """
        managers.pinned_apps.PinAppID(self.app_id) # type: ignore

    def unpin_app(self):
        """
        Unpin this window's app (corresponds to the 'show windows from this app on all desktops' toggle).
        """
        managers.pinned_apps.UnpinAppID(self.app_id) # type: ignore

    def is_app_pinned(self) -> bool:
        """
        Check if this window's app is pinned (corresponds to the 'show windows from this app on all desktops' toggle).

        Returns:
            bool: is the app pinned?.
        """
        return managers.pinned_apps.IsAppIdPinned(self.app_id) # type: ignore


    #  ------------------------------------------------
    #  IVirtualDesktopManagerInternal methods
    #  ------------------------------------------------
    def move(self, desktop: VirtualDesktop):
        """Move the window to a different virtual desktop.

        Args:
            desktop (VirtualDesktop): Desktop to move the window to.

        Example:

                >>> AppView.current().move_to_desktop(VirtualDesktop(1))

        """
        managers.manager_internal.MoveViewToDesktop(self._view, desktop._virtual_desktop)  # type: ignore

    @property
    def desktop_id(self) -> GUID:
        """
        Returns:
            GUID -- The ID of the desktop which the window is on.
        """
        return self._view.GetVirtualDesktopId() # type: ignore

    @property
    def desktop(self) -> VirtualDesktop:
        """
        Returns:
            VirtualDesktop: The virtual desktop which this window is on.
        """
        return VirtualDesktop(desktop_id=self.desktop_id)


    def is_on_desktop(self, desktop: VirtualDesktop, include_pinned: bool = True) -> bool:
        """Is this window on the passed virtual desktop?

        Args:
            desktop (VirtualDesktop): Desktop to check
            include_pinned (bool, optional): Also return `True` for pinned windows

        Example:

            >>> AppView.current().is_on_desktop(VirtualDesktop(1))

        """
        if include_pinned:
            return (self.desktop_id == desktop.id) or self.is_pinned() or self.is_app_pinned()
        else:
            return self.desktop_id == desktop.id


    def is_on_current_desktop(self) -> bool:
        """Is this window on the current desktop?
        """
        return self.is_on_desktop(VirtualDesktop.current())


def get_apps_by_z_order(switcher_windows: bool = True, current_desktop: bool = True) -> List[AppView]:
    """Get a list of AppViews, ordered by their Z position, with
    the foreground window first.

    Args:
        switcher_windows (bool, optional): Only include windows which appear in the alt-tab dialogue. Defaults to True.
        current_desktop (bool, optional): Only include windows which are on the current virtual desktop. Defaults to True.

    Returns:
        List[AppView]: AppViews matching the specified criteria.
    """
    views_arr = managers.view_collection.GetViewsByZOrder() # type: ignore
    all_views = [AppView(view=v) for v in views_arr.iter(IApplicationView)]
    if not switcher_windows and not current_desktop:
        # no filters
        return all_views
    else:
        result = []
        vd = VirtualDesktop.current()
        for view in all_views:
            if switcher_windows and not view.is_shown_in_switchers():
                continue
            if current_desktop and not view.is_on_desktop(vd):
                continue
            result.append(view)
        return result


class VirtualDesktop():
    """
    Wrapper around the `IVirtualDesktop` COM object, representing one virtual desktop.
    """
    def __init__(
        self,
        number: Optional[int] = None,
        desktop_id: Optional[GUID] = None,
        desktop: Optional['IVirtualDesktop'] = None,
        current: Optional[bool] = False
    ):
        """One of the following arguments must be provided:

        Args:
            number (int, optional): The number of the desired desktop in the task view (1-indexed). Defaults to None.
            desktop_id (GUID, optional): A desktop GUID. Defaults to None.
            desktop (IVirtualDesktop, optional): An `IVirtualDesktop`. Defaults to None.
            current (bool, optional): The current virtual desktop. Defaults to False.
        """

        if number:
            if number <= 0:
                raise ValueError(f"Desktop number must be at least 1, {number} provided")
            array = managers.manager_internal.get_all_desktops() # type: ignore
            desktop_count = array.GetCount()
            if number > desktop_count:
                raise ValueError(
                    f"Desktop number {number} exceeds the number of desktops, {desktop_count}."
                )
            self._virtual_desktop = array.get_at(number - 1, IVirtualDesktop)

        elif desktop_id:
            self._virtual_desktop = managers.manager_internal.FindDesktop(desktop_id) # type: ignore

        elif desktop:
            self._virtual_desktop = desktop

        elif current:
            self._virtual_desktop = managers.manager_internal.get_current_desktop() # type: ignore

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

    @classmethod
    def create(cls):
        """Create a new virtual desktop.

        Returns:
            VirtualDesktop: The created desktop.
        """
        desktop = managers.manager_internal.create_desktop() # type: ignore
        return cls(desktop=desktop)

    @property
    def id(self) -> GUID:
        """The GUID of this desktop.

        Returns:
            GUID: The unique id for this desktop.
        """
        return self._virtual_desktop.GetID() # type: ignore

    @property
    def number(self) -> int:
        """The index of this virtual desktop in the task view. Between 1 and
        the total number of desktops active.

        Returns:
            int: The desktop number.
        """
        array = managers.manager_internal.get_all_desktops() # type: ignore
        for i, vd in enumerate(array.iter(IVirtualDesktop), 1):
            if self.id == vd.GetID():
                return i
        else:
            raise Exception(f"Desktop with ID {self.id} not found")

    @property
    def name(self) -> str:
        """The name of this virtual desktop in the task view.
        Note that the default name is an empty string even though the task view shows
        e.g. 'Desktop 4'.

        Returns:
            str: The desktop name.

        Raises:
            NotImplementedError: If the Windows version is < 19041.
        """
        if build.OVER_21313:
            return str(self._virtual_desktop.GetName()) # type: ignore

        if not build.OVER_19041:
            raise NotImplementedError(f"{VirtualDesktop.name.fget.__name__} is not supported on < 19041 versions")

        array = managers.manager_internal.get_all_desktops() # type: ignore
        for vd in array.iter(IVirtualDesktop2):
            if self.id == vd.GetID():
                return str(vd.GetName())
        else:
            raise Exception(f"Desktop with ID {self.id} not found")

    def rename(self, name: str):
        """Rename this desktop.

        Args:
            name: The new name for this desktop.

        Raises:
            NotImplementedError: If the Windows version is < 19041.
        """

        if managers.manager_internal2 is not None: # type: ignore
            managers.manager_internal2.SetName(self._virtual_desktop, HSTRING(name)) # type: ignore
            return

        if not build.OVER_19041:
            raise NotImplementedError(f"{VirtualDesktop.rename.__name__} is not supported on < 19041 versions")

        managers.manager_internal.SetName(self._virtual_desktop, HSTRING(name)) # type: ignore

    def remove(self, fallback: Optional[VirtualDesktop] = None):
        """Delete this virtual desktop, falling back to 'fallback'.

        Args:
            fallback (VirtualDesktop, optional): If you are currently on the desktop
            you pass to this method, focus will be shifted to the desktop passed here.
            If no desktop is passed, it will default to the first.
        """
        if fallback is None:
            fallback = VirtualDesktop(1)
        managers.manager_internal.RemoveDesktop(self._virtual_desktop, fallback._virtual_desktop) # type: ignore

    def go(self, allow_set_foreground: bool = True):
        """Switch to this virtual desktop.

        Args:
            allow_set_foreground (bool, optional): Call AllowSetForegroundWindow(ASFW_ANY) before switching. This partially fixes an issue where the focus remains behind after switching. Defaults to True.

        Note:
            More details at https://github.com/Ciantic/VirtualDesktopAccessor/issues/4 and https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-allowsetforegroundwindow.
        """
        if allow_set_foreground:
            windll.user32.AllowSetForegroundWindow(ASFW_ANY)
        managers.manager_internal.switch_desktop(self._virtual_desktop) # type: ignore

    def apps_by_z_order(self, include_pinned: bool = True) -> List[AppView]:
        """Get a list of AppViews, ordered by their Z position, with
        the foreground window first.

        Args:
            switcher_windows (bool, optional): Only include windows which appear in the alt-tab dialogue. Defaults to True.
            current_desktop (bool, optional): Only include windows which are on the current virtual desktop. Defaults to True.

        Returns:
            List[AppView]: AppViews matching the specified criteria.
        """
        views_arr = managers.view_collection.GetViewsByZOrder() # type: ignore
        all_views = [AppView(view=v) for v in views_arr.iter(IApplicationView)]
        result = []
        for view in all_views:
            if view.is_shown_in_switchers() and view.is_on_desktop(self, include_pinned):
                result.append(view)
        return result

    def set_wallpaper(self, path: str):
        """Set wallpaper on current virtual desktop to `path`.

        Args:
            path (str): path to wallpaper file
        """
        if build.OVER_21313:
            managers.manager_internal.SetWallpaper(self._virtual_desktop,path=HSTRING(path)) # type: ignore
        else:
            raise NotImplementedError("set_wallpaper is only available on Windows 11")


def get_virtual_desktops() -> List[VirtualDesktop]:
    """Return a list of all current virtual desktops, one for each desktop visible in the task view.

    Returns:
        List[VirtualDesktop]: Virtual desktops currently active.
    """
    array = managers.manager_internal.get_all_desktops() # type: ignore
    return [VirtualDesktop(desktop=vd) for vd in array.iter(IVirtualDesktop)]


def set_wallpaper_for_all_desktops(path: str):
    """Set wallpaper on current virtual desktop to `path`.

    Args:
        path (str): path to wallpaper file
    """
    if build.OVER_21313:
        managers.manager_internal.SetWallpaperForAllDesktops(path=HSTRING(path)) # type: ignore
    else:
        raise NotImplementedError("set_wallpaper_for_all_desktops is only available on Windows 11")
