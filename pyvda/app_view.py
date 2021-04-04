from typing import List

from comtypes import GUID
from .desktop import VirtualDesktop
from .win10desktops import IApplicationView
from .utils import (
    get_vd_manager_internal,
    get_view_collection,
    get_pinned_apps,
)

class AppView():
    """
    A wrapper around an `IApplicationView` object exposing window functionality relating to:

        * Setting focus
        * Pinning and unpinning (making a window persistent across all virtual desktops)
        * Moving a window between virtual desktops

    """

    def __init__(self, hwnd: int = None, view: 'IApplicationView' = None):
        """One of the following parameters must be provided:

        Args:
            hwnd (int, optional): Handle to a window. Defaults to None.
            view (IApplicationView, optional): An `IApplicationView` object. Defaults to None.
        """
        if hwnd:
            # Get the IApplicationView for the window
            view_collection = get_view_collection()
            self._view = view_collection.GetViewForHwnd(hwnd)
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
        return self._view.GetThumbnailWindow()

    @property
    def app_id(self) -> int:
        """The ID of this window's app.
        """
        return self._view.GetAppUserModelId()

    @classmethod
    def current(cls):
        """
        Returns:
            AppView: An AppView for the currently focused window.
        """
        view_collection = get_view_collection()
        focused = view_collection.GetViewInFocus()
        return cls(view=focused)

    #  ------------------------------------------------
    #  IApplicationView methods
    #  ------------------------------------------------
    def is_shown_in_switchers(self) -> bool:
        """Is the view shown in the alt-tab view?
        """
        return bool(self._view.GetShowInSwitchers())

    def is_visible(self) -> bool:
        """Is the view visible?
        """
        return bool(self._view.GetVisibility())

    def get_activation_timestamp(self) -> int:
        """Get the last activation timestamp for this window.
        """
        return self._view.GetLastActivationTimestamp()

    def set_focus(self):
        """Focus the window"""
        return self._view.SetFocus()

    def switch_to(self):
        """Switch to the window. Behaves slightly differently to set_focus -
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
        Check if this window is pinned (corresponds to the 'show window on all desktops' toggle).

        Returns:
            bool: is the window pinned?
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
            bool: is the app pinned?.
        """
        pinnedApps = get_pinned_apps()
        return pinnedApps.IsAppIdPinned(self.app_id)


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
        manager_internal = get_vd_manager_internal()
        manager_internal.MoveViewToDesktop(self._view, desktop._virtual_desktop)

    @property
    def desktop_id(self) -> GUID:
        """
        Returns:
            GUID -- The ID of the desktop which the window is on.
        """
        return self._view.GetVirtualDesktopId()

    @property
    def desktop(self) -> VirtualDesktop:
        """
        Returns:
            VirtualDesktop: The virtual desktop which this window is on.
        """
        return VirtualDesktop(desktop_id=self.desktop_id)


    def is_on_desktop(self, desktop: VirtualDesktop) -> bool:
        """Is this window on the passed virtual desktop?

        Args:
            desktop (VirtualDesktop): Desktop to check

        Example:

            >>> AppView.current().is_on_desktop(VirtualDesktop(1))

        """
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
    collection = get_view_collection()
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

