"""
References:
    * https://github.com/Ciantic/VirtualDesktopAccessor/blob/master/VirtualDesktopAccessor/dllmain.h
"""
from ctypes import POINTER
from typing import List

from comtypes.GUID import GUID
from .desktop import VirtualDesktop
from .win10desktops import (
    IVirtualDesktop,
    IApplicationView,
)
from .utils import (
    get_vd_manager,
    get_vd_manager_internal,
    get_view_collection,
    get_pinned_apps,
)

class AppView():
    def __init__(self, hwnd: int = None, view: IApplicationView = None):
        """
        # TODO:
        """
        if hwnd:
            # Get the IApplicationView for the window
            view_collection = get_view_collection()
            self._view = view_collection.GetViewForHwnd(hwnd)
        elif view:
            self._view = view
        else:
            raise Exception(f"Must pass 'hwnd' or 'view'")

    @property
    def hwnd(self) -> int:
        """
        # TODO:
        """
        return self._view.GetThumbnailWindow()

    @property
    def app_id(self) -> int:
        """
        # TODO:
        """
        return self._view.GetAppUserModelId()

    @classmethod
    def current(cls):
        """
        # TODO:
        """
        view_collection = get_view_collection()
        focused = view_collection.GetViewInFocus()
        return cls(view=focused)

    #  ------------------------------------------------
    #  IApplicationView methods
    #  ------------------------------------------------
    def is_shown_in_switchers(self) -> bool:
        """
        Returns:
            bool -- is the view shown in the alt-tab view?
        """
        return bool(self._view.GetShowInSwitchers())


    def is_visible(self) -> bool:
        """
        Returns:
            bool -- is the view visible?
        """
        return bool(self._view.GetVisibility())


    def last_activation_timestamp(self) -> int:
        """Get the last activation for this window.

        Returns:
            int -- last activation timestamp
        """
        return self._view.GetLastActivationTimestamp()

    def set_focus(self):
        """Focus the window"""
        return self._view.SetFocus()

    def switch_to(self):
        """Switch to the window. Behaves slightly differently to ViewSetFocus -
        this is what is called when you use the alt-tab menu."""
        return self._view.SwitchTo()


    #  ------------------------------------------------
    #  IVirtualDesktopPinnedApps methods
    #  ------------------------------------------------
    def pin(self):
        """
        Pin the window (corresponds to the 'show window on all desktops' toggle).
        """
        pinnedApps = get_pinned_apps()
        pinnedApps.PinView(self._view)

    def unpin(self):
        """
        Unpin the window (corresponds to the 'show window on all desktops' toggle).
        """
        pinnedApps = get_pinned_apps()
        pinnedApps.UnpinView(self._view)

    def is_pinned(self) -> bool:
        """
        Check if a window is pinned (corresponds to the 'show window on all desktops' toggle).

        Returns:
            bool -- is the window pinned?.
        """
        pinnedApps = get_pinned_apps()
        return pinnedApps.IsViewPinned(self._view)

    def pin_app(self):
        """
        Pin this window's app (corresponds to the 'show windows from this app on all desktops' toggle).
        """
        pinnedApps = get_pinned_apps()
        pinnedApps.PinAppID(self.app_id)

    def unpin_app(self):
        """
        Unpin this window's app (corresponds to the 'show windows from this app on all desktops' toggle).
        """
        pinnedApps = get_pinned_apps()
        pinnedApps.UnpinAppID(self.app_id)

    def is_app_pinned(self) -> bool:
        """
        Check if this window's app is pinned (corresponds to the 'show windows from this app on all desktops' toggle).

        Returns:
            bool -- is the app pinned?.
        """
        pinnedApps = get_pinned_apps()
        return pinnedApps.IsAppIdPinned(self.app_id)


    #  ------------------------------------------------
    #  IVirtualDesktopManagerInternal methods
    #  ------------------------------------------------
    def move_to_desktop(self, desktop: VirtualDesktop):
        """
        Move the window to a different virtual desktop.

        # TODO:
        Arguments:
            number {int} -- Desktop number to move the window to, between 1 and the number of desktops active.
        """
        manager_internal = get_vd_manager_internal()
        manager_internal.MoveViewToDesktop(self._view, desktop._virtual_desktop)

    @property
    def desktop_id(self) -> GUID:
        """
        # TODO:
        Returns the number of the desktop which the window is on.

        Returns:
            int -- Its desktop number.
        """
        return self._view.GetVirtualDesktopId()

    @property
    def desktop(self) -> VirtualDesktop:
        """
        # TODO:
        Returns the number of the desktop which the window is on.

        Returns:
            int -- Its desktop number.
        """
        return VirtualDesktop(desktop_id=self.desktop_id)

    def is_on_desktop(self, desktop: VirtualDesktop) -> bool:
        """
        # TODO:
        """
        return self.desktop_id == desktop.id

    def is_on_current_desktop(self) -> bool:
        """
        # TODO:
        """
        return self.is_on_desktop(VirtualDesktop.current())


def get_apps_by_z_order(switcher_windows: bool = True, current_desktop: bool = True) -> List[AppView]:
    """Get a list of window handles, ordered by their Z position, with
    the foreground window first.

    Arguments:
        switcher_windows {bool} -- Only include windows which appear in the alt-tab dialogue
        current_desktop {bool} -- Only include windows which are on the current virtual desktop

    Returns:
        List[int] -- Window handles
    """
    collection = _get_view_collection()
    views_arr = collection.GetViewsByZOrder()
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

