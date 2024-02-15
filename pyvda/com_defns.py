"""
References:
    * GUIDs, Interfaces: https://github.com/Ciantic/VirtualDesktopAccessor/blob/master/VirtualDesktopAccessor/Win10Desktops.h
    * Service provider: https://docs.microsoft.com/en-us/dotnet/api/system.iserviceprovider?view=netframework-4.8
    * Object array: https://docs.microsoft.com/en-us/windows/win32/api/objectarray/nn-objectarray-iobjectarray
    * HSTRING: https://docs.microsoft.com/en-us/uwp/cpp-ref-for-winrt/hstring
    * comtypes: http://svn.python.org/projects/ctypes/tags/comtypes-0.3.2/docs/com_interfaces.html
    * dev blogs:
        * http://grabacr.net/archives/5601
        * https://www.cyberforum.ru/blogs/105416/blog3671.html
"""
import sys
import os
from ctypes import c_ulonglong, POINTER, Structure, HRESULT
from ctypes.wintypes import (
    UINT,
    INT,
    WCHAR,
    LPVOID,
    ULONG,
    HWND,
    RECT,
    LPCWSTR,
    BOOL,
    SIZE,
    DWORD,
)
from typing import Any, Iterator
from comtypes import (
    IUnknown,
    GUID,
    STDMETHOD,
    COMMETHOD,
)
from .winstring import HSTRING

if not os.getenv("READTHEDOCS"):
    winver = sys.getwindowsversion()
    major, minor, build = winver.platform_version or winver[:3]
    BUILD_OVER_20231 = build >= 20231
    BUILD_OVER_21313 = build >= 21313
    BUILD_OVER_22449 = build >= 22449
    BUILD_OVER_22621 = build >= 22621
    BUILD_OVER_22631 = build >= 22631
else:
    BUILD_OVER_20231 = False
    BUILD_OVER_21313 = False
    BUILD_OVER_22449 = False
    BUILD_OVER_22621 = False
    BUILD_OVER_22631 = False


CLSID_ImmersiveShell = GUID("{C2F03A33-21F5-47FA-B4BB-156362A2F239}")
CLSID_VirtualDesktopManagerInternal = GUID("{C5E0CDCA-7B6E-41B2-9FC4-D93975CC467B}")
CLSID_IVirtualDesktopManager = GUID("{AA509086-5CA9-4C25-8F95-589D3C07B48A}")
CLSID_VirtualDesktopPinnedApps = GUID("{B5A399E7-1C87-46B8-88E9-FC5747B171BD}")

# Ignore following APIs:
IAsyncCallback = UINT
IImmersiveMonitor = UINT
APPLICATION_VIEW_COMPATIBILITY_POLICY = UINT
IShellPositionerPriority = UINT
IApplicationViewOperation = UINT
APPLICATION_VIEW_CLOAK_TYPE = UINT
IApplicationViewPosition = UINT
IImmersiveApplication = UINT
IApplicationViewChangeListener = UINT

TrustLevel = INT
AdjacentDesktop = UINT

ULONGLONG = c_ulonglong
PWSTR = POINTER(WCHAR)
REFGUID = POINTER(GUID)
REFIID = POINTER(GUID)


class IServiceProvider(IUnknown):
    _iid_ = GUID("{6D5140C1-7436-11CE-8034-00AA006009FA}")
    _methods_ = [
        STDMETHOD(HRESULT, "QueryService", (REFGUID, REFIID, POINTER(LPVOID),)),
    ]


class IObjectArray(IUnknown):
    _iid_ = GUID("{92CA9DCD-5622-4BBA-A805-5E9F541BD8C9}")
    _methods_ = [
        COMMETHOD([], HRESULT, "GetCount", (["out"], POINTER(UINT), "pcObjects"),),
        STDMETHOD(HRESULT, "GetAt", (UINT, REFIID, POINTER(LPVOID),)),
    ]

    def get_at(self, i: int, cls: Any) -> Any:
        item = POINTER(cls)()
        self.GetAt(i, cls._iid_, item)
        return item

    def iter(self, cls: Any) -> Iterator[Any]:
        for i in range(self.GetCount()):
            yield self.get_at(i, cls)


# Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Interface\{372E1D3B-38D3-42E4-A15B-8AB2B178F513}
# Found with searching "IApplicationView"
class IApplicationView(IUnknown):
    _iid_ = GUID("{372E1D3B-38D3-42E4-A15B-8AB2B178F513}")

IApplicationView._methods_ = [
    # IInspectable methods
    STDMETHOD(HRESULT, "GetIids", (POINTER(ULONG), POINTER(POINTER(GUID)),)),
    STDMETHOD(HRESULT, "GetRuntimeClassName", (POINTER(HSTRING),)),
    STDMETHOD(HRESULT, "GetTrustLevel", (POINTER(TrustLevel),)),
    # IApplicationView methods
    STDMETHOD(HRESULT, "SetFocus", ()),
    STDMETHOD(HRESULT, "SwitchTo", ()),
    STDMETHOD(HRESULT, "TryInvokeBack", (POINTER(IAsyncCallback),)),
    COMMETHOD([], HRESULT, "GetThumbnailWindow", (["out"], POINTER(HWND), "hwnd")),
    STDMETHOD(HRESULT, "GetMonitor", (POINTER(POINTER(IImmersiveMonitor)),)),
    COMMETHOD([], HRESULT, "GetVisibility", (["out"], POINTER(UINT), "pVisible")),
    STDMETHOD(HRESULT, "SetCloak", (APPLICATION_VIEW_CLOAK_TYPE, UINT,)),
    STDMETHOD(HRESULT, "GetPosition", (REFIID, POINTER(LPVOID),)),
    STDMETHOD(HRESULT, "SetPosition", (POINTER(IApplicationViewPosition),)),
    STDMETHOD(HRESULT, "InsertAfterWindow", (HWND,)),
    STDMETHOD(HRESULT, "GetExtendedFramePosition", (POINTER(RECT),)),
    COMMETHOD([], HRESULT, "GetAppUserModelId", (["out"], POINTER(PWSTR), "pId")),
    STDMETHOD(HRESULT, "SetAppUserModelId", (LPCWSTR,)),
    STDMETHOD(HRESULT, "IsEqualByAppUserModelId", (LPCWSTR, POINTER(UINT),)),
    STDMETHOD(HRESULT, "GetViewState", (POINTER(UINT),)),
    STDMETHOD(HRESULT, "SetViewState", (UINT,)),
    STDMETHOD(HRESULT, "GetNeediness", (POINTER(UINT),)),
    COMMETHOD([], HRESULT, "GetLastActivationTimestamp", (["out"], POINTER(ULONGLONG), "pGuid")),
    STDMETHOD(HRESULT, "SetLastActivationTimestamp", (ULONGLONG,)),
    COMMETHOD([], HRESULT, "GetVirtualDesktopId", (["out"], POINTER(GUID), "pGuid")),
    STDMETHOD(HRESULT, "SetVirtualDesktopId", (REFGUID,)),
    COMMETHOD([], HRESULT, "GetShowInSwitchers", (["out"], POINTER(UINT), "pShown")),
    STDMETHOD(HRESULT, "SetShowInSwitchers", (UINT,)),
    STDMETHOD(HRESULT, "GetScaleFactor", (POINTER(UINT),)),
    STDMETHOD(HRESULT, "CanReceiveInput", (POINTER(BOOL),)),
    STDMETHOD(HRESULT, "GetCompatibilityPolicyType", (POINTER(APPLICATION_VIEW_COMPATIBILITY_POLICY),)),
    STDMETHOD(HRESULT, "SetCompatibilityPolicyType", (APPLICATION_VIEW_COMPATIBILITY_POLICY,)),
    # STDMETHOD(HRESULT, "GetPositionPriority", (POINTER(POINTER(IShellPositionerPriority)),)),
    # STDMETHOD(HRESULT, "SetPositionPriority", (POINTER(IShellPositionerPriority),)),
    STDMETHOD(HRESULT, "GetSizeConstraints", (POINTER(IImmersiveMonitor), POINTER(SIZE), POINTER(SIZE))),
    STDMETHOD(HRESULT, "GetSizeConstraintsForDpi", (UINT, POINTER(SIZE), POINTER(SIZE),)),
    STDMETHOD(HRESULT, "SetSizeConstraintsForDpi", (POINTER(UINT), POINTER(SIZE), POINTER(SIZE))),
    # STDMETHOD(HRESULT, "QuerySizeConstraintsFromApp", ()),
    STDMETHOD(HRESULT, "OnMinSizePreferencesUpdated", (HWND,)),
    STDMETHOD(HRESULT, "ApplyOperation", (POINTER(IApplicationViewOperation),)),
    STDMETHOD(HRESULT, "IsTray", (POINTER(BOOL),)),
    STDMETHOD(HRESULT, "IsInHighZOrderBand", (POINTER(BOOL),)),
    STDMETHOD(HRESULT, "IsSplashScreenPresented", (POINTER(BOOL),)),
    STDMETHOD(HRESULT, "Flash", ()),
    STDMETHOD(HRESULT, "GetRootSwitchableOwner", (POINTER(POINTER(IApplicationView)),)),
    STDMETHOD(HRESULT, "EnumerateOwnershipTree", (POINTER(POINTER(IObjectArray)),)),
    STDMETHOD(HRESULT, "GetEnterpriseId", (POINTER(PWSTR),)),
    STDMETHOD(HRESULT, "IsMirrored", (POINTER(BOOL),)),
    STDMETHOD(HRESULT, "GetFrameworkViewType", (POINTER(UINT),)),
    STDMETHOD(HRESULT, "GetCanTab", (POINTER(UINT),)),
    STDMETHOD(HRESULT, "SetCanTab", (UINT,)),
    STDMETHOD(HRESULT, "GetIsTabbed", (POINTER(UINT),)),
    STDMETHOD(HRESULT, "SetIsTabbed", (UINT,)),
    STDMETHOD(HRESULT, "RefreshCanTab", ()),
    STDMETHOD(HRESULT, "GetIsOccluded", (POINTER(UINT),)),
    STDMETHOD(HRESULT, "SetIsOccluded", (UINT,)),
    STDMETHOD(HRESULT, "UpdateEngagementFlags", (UINT, UINT,)),
    STDMETHOD(HRESULT, "SetForceActiveWindowAppearance", (UINT,)),
    STDMETHOD(HRESULT, "GetLastActivationFILETIME", (POINTER(SIZE),)),
    STDMETHOD(HRESULT, "GetPersistingStateName", (POINTER(PWSTR),)),
]

# No change for 22449
if BUILD_OVER_22631:
    GUID_IVirtualDesktop = GUID("{3F07F4BE-B107-441A-AF0F-39D82529072C}")
elif BUILD_OVER_22621:
    GUID_IVirtualDesktop = GUID("{3F07F4BE-B107-441A-AF0F-39D82529072C}")
elif BUILD_OVER_21313:
    GUID_IVirtualDesktop = GUID("{536D3495-B208-4CC9-AE26-DE8111275BF8}")
elif BUILD_OVER_20231:
    GUID_IVirtualDesktop = GUID("{62FDF88B-11CA-4AFB-8BD8-2296DFAE49E2}")
else:
    GUID_IVirtualDesktop = GUID("{FF72FFDD-BE7E-43FC-9C03-AD81681E88E4}")

# In registry: Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Interface\{FF72FFDD-BE7E-43FC-9C03-AD81681E88E4}
class IVirtualDesktop(IUnknown):
    _iid_ = GUID_IVirtualDesktop
    if BUILD_OVER_22621:
        _methods_ = [
            STDMETHOD(HRESULT, "IsViewVisible", (POINTER(IApplicationView), POINTER(UINT))),
            COMMETHOD([], HRESULT, "GetID", (["out"], POINTER(GUID), "pGuid"),),
            COMMETHOD([], HRESULT, "GetName", (["out"], POINTER(HSTRING), "pName"),),
            COMMETHOD([], HRESULT, "GetWallpaperPath", (["out"], POINTER(HSTRING), "pPath"),),
            COMMETHOD([], HRESULT, "IsRemote", (["out"], POINTER(HWND), "pW"), ),
        ]
    elif BUILD_OVER_21313:
        _methods_ = [
            STDMETHOD(HRESULT, "IsViewVisible", (POINTER(IApplicationView), POINTER(UINT))),
            COMMETHOD([], HRESULT, "GetID", (["out"], POINTER(GUID), "pGuid"),),
            COMMETHOD([], HRESULT, "IsRemote", (["out"], POINTER(HWND), "pW"),),
            COMMETHOD([], HRESULT, "GetName", (["out"], POINTER(HSTRING), "pName"),),
            COMMETHOD([], HRESULT, "GetWallpaperPath", (["out"], POINTER(HSTRING), "pPath"),),
        ]
    else:
        _methods_ = [
            STDMETHOD(HRESULT, "IsViewVisible", (POINTER(IApplicationView), POINTER(UINT))),
            COMMETHOD([], HRESULT, "GetID", (["out"], POINTER(GUID), "pGuid"),),
        ]


GUID_IVirtualDesktop2 = GUID("{31EBDE3F-6EC3-4CBD-B9FB-0EF6D09B41F4}")
class IVirtualDesktop2(IUnknown):
    _iid_ = GUID_IVirtualDesktop2
    _methods_ = [
        STDMETHOD(HRESULT, "IsViewVisible", (POINTER(IApplicationView), POINTER(UINT))),
        COMMETHOD([], HRESULT, "GetID", (["out"], POINTER(GUID), "pGuid"),),
        COMMETHOD([], HRESULT, "GetName", (["out"], POINTER(HSTRING), "pName"),),
    ]


# Same GUID for 22449
if BUILD_OVER_22631:
    GUID_IVirtualDesktopManagerInternal = GUID("{4970BA3D-FD4E-4647-BEA3-D89076EF4B9C}")
elif BUILD_OVER_22621:
    GUID_IVirtualDesktopManagerInternal = GUID("{A3175F2D-239C-4BD2-8AA0-EEBA8B0B138E}")
elif BUILD_OVER_21313:
    GUID_IVirtualDesktopManagerInternal = GUID("{B2F925B9-5A0F-4D2E-9F4D-2B1507593C10}")
elif BUILD_OVER_20231:
    GUID_IVirtualDesktopManagerInternal = GUID("{094AFE11-44F2-4BA0-976F-29A97E263EE0}")
else:
    GUID_IVirtualDesktopManagerInternal = GUID("{F31574D6-B682-4CDC-BD56-1827860ABEC6}")

# HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Interface\{F31574D6-B682-4CDC-BD56-1827860ABEC6}
class IVirtualDesktopManagerInternal(IUnknown):
    _iid_ = GUID_IVirtualDesktopManagerInternal
    if BUILD_OVER_22631:
        _methods_ = [
            COMMETHOD([], HRESULT, "GetCount",  (["out"], POINTER(UINT), "pCount"),),
            STDMETHOD(HRESULT, "MoveViewToDesktop", (POINTER(IApplicationView), POINTER(IVirtualDesktop))),
            STDMETHOD(HRESULT, "CanViewMoveDesktops", (POINTER(IApplicationView), POINTER(UINT))),
            COMMETHOD([], HRESULT, "GetCurrentDesktop", (["out"], POINTER(POINTER(IVirtualDesktop)), "pDesktop"),),
            COMMETHOD([], HRESULT, "GetDesktops", (["out"], POINTER(POINTER(IObjectArray)), "array")),
            STDMETHOD(HRESULT, "GetAdjacentDesktop", (POINTER(IVirtualDesktop), AdjacentDesktop, POINTER(POINTER(IVirtualDesktop)),)),
            STDMETHOD(HRESULT, "SwitchDesktop", (POINTER(IVirtualDesktop),)),
            STDMETHOD(HRESULT, "Unknown1", (POINTER(IVirtualDesktop),)),
            COMMETHOD([], HRESULT, "CreateDesktopW", (["out"], POINTER(POINTER(IVirtualDesktop)), "pDesktop"),),
            STDMETHOD(HRESULT, "MoveDesktop", (POINTER(IVirtualDesktop), HWND, INT)),
            COMMETHOD([], HRESULT, "RemoveDesktop", (["in"], POINTER(IVirtualDesktop), "destroyDesktop"), (["in"], POINTER(IVirtualDesktop), "fallbackDesktop")),
            COMMETHOD([], HRESULT, "FindDesktop", (["in"], POINTER(GUID), "pGuid"), (["out"], POINTER(POINTER(IVirtualDesktop)), "pDesktop")),
            STDMETHOD(HRESULT, "Unknown2", (POINTER(IVirtualDesktop), POINTER(POINTER(IObjectArray)), POINTER(POINTER(IObjectArray)))),
            COMMETHOD([], HRESULT, "SetName", (["in"], POINTER(IVirtualDesktop), "pDesktop"), (["in"], HSTRING, "name")),
            COMMETHOD([], HRESULT, "SetWallpaper", (["in"], POINTER(IVirtualDesktop), "pDesktop"), (["in"], HSTRING, "path")),
            COMMETHOD([], HRESULT, "SetWallpaperForAllDesktops", (["in"], HSTRING, "path")),
            COMMETHOD([], HRESULT, "CopyDesktopState", (["in"], POINTER(IApplicationView), "pView0"), (["in"], POINTER(IApplicationView), "pView0")),

            COMMETHOD([], HRESULT, "Unknown3", (["in"], HSTRING, "a1"), (["out"], POINTER(POINTER(IVirtualDesktop)), "out")),
            STDMETHOD(HRESULT, "pDesktop", (POINTER(IVirtualDesktop),)),
            STDMETHOD(HRESULT, "Unknown5", (POINTER(IVirtualDesktop),)),
            COMMETHOD([], HRESULT, "Unknown6", (["out"], POINTER(POINTER(IVirtualDesktop)), "pDesktop"),),
            STDMETHOD(HRESULT, "Unknown7"),
        ]
    if BUILD_OVER_22621:
        _methods_ = [
            COMMETHOD([], HRESULT, "GetCount",  (["out"], POINTER(UINT), "pCount"),),
            STDMETHOD(HRESULT, "MoveViewToDesktop", (POINTER(IApplicationView), POINTER(IVirtualDesktop))),
            STDMETHOD(HRESULT, "CanViewMoveDesktops", (POINTER(IApplicationView), POINTER(UINT))),
            COMMETHOD([], HRESULT, "GetCurrentDesktop", (["out"], POINTER(POINTER(IVirtualDesktop)), "pDesktop"),),
            COMMETHOD([], HRESULT, "GetDesktops", (["out"], POINTER(POINTER(IObjectArray)), "array")),
            STDMETHOD(HRESULT, "GetAdjacentDesktop", (POINTER(IVirtualDesktop), AdjacentDesktop, POINTER(POINTER(IVirtualDesktop)),)),
            STDMETHOD(HRESULT, "SwitchDesktop", (POINTER(IVirtualDesktop),)),
            COMMETHOD([], HRESULT, "CreateDesktopW", (["out"], POINTER(POINTER(IVirtualDesktop)), "pDesktop"),),
            STDMETHOD(HRESULT, "MoveDesktop", (POINTER(IVirtualDesktop), HWND, INT)),
            COMMETHOD([], HRESULT, "RemoveDesktop", (["in"], POINTER(IVirtualDesktop), "destroyDesktop"), (["in"], POINTER(IVirtualDesktop), "fallbackDesktop")),
            COMMETHOD([], HRESULT, "FindDesktop", (["in"], POINTER(GUID), "pGuid"), (["out"], POINTER(POINTER(IVirtualDesktop)), "pDesktop")),
            STDMETHOD(HRESULT, "Unknown", (POINTER(IVirtualDesktop), POINTER(POINTER(IObjectArray)), POINTER(POINTER(IObjectArray)))),
            COMMETHOD([], HRESULT, "SetName", (["in"], POINTER(IVirtualDesktop), "pDesktop"), (["in"], HSTRING, "name")),
            COMMETHOD([], HRESULT, "SetWallpaper", (["in"], POINTER(IVirtualDesktop), "pDesktop"), (["in"], HSTRING, "path")),
            COMMETHOD([], HRESULT, "SetWallpaperForAllDesktops", (["in"], HSTRING, "path")),
            COMMETHOD([], HRESULT, "CopyDesktopState", (["in"], POINTER(IApplicationView), "pView0"), (["in"], POINTER(IApplicationView), "pView0")),
            COMMETHOD([], HRESULT, "GetDesktopPerMonitor", (["out"], POINTER(BOOL), "state")),
            COMMETHOD([], HRESULT, "SetDesktopPerMonitor", (["in"], BOOL, "state")),
        ]
    elif BUILD_OVER_22449:
        _methods_ = [
            COMMETHOD([], HRESULT, "GetCount", (["in"], HWND, "hwnd"), (["out"], POINTER(UINT), "pCount"),),
            STDMETHOD(HRESULT, "MoveViewToDesktop", (POINTER(IApplicationView), POINTER(IVirtualDesktop))),
            STDMETHOD(HRESULT, "CanViewMoveDesktops", (POINTER(IApplicationView), POINTER(UINT))),
            COMMETHOD([], HRESULT, "GetCurrentDesktop", (["in"], HWND, "hwnd"), (["out"], POINTER(POINTER(IVirtualDesktop)), "pDesktop"),),
            COMMETHOD([], HRESULT, "GetAllCurrentDesktops", (["out"], POINTER(POINTER(IObjectArray)), "array")),
            COMMETHOD([], HRESULT, "GetDesktops", (["in"], HWND, "hwnd"), (["out"], POINTER(POINTER(IObjectArray)), "array")),
            STDMETHOD(HRESULT, "GetAdjacentDesktop", (POINTER(IVirtualDesktop), AdjacentDesktop, POINTER(POINTER(IVirtualDesktop)),)),
            STDMETHOD(HRESULT, "SwitchDesktop", (HWND, POINTER(IVirtualDesktop),)),
            COMMETHOD([], HRESULT, "CreateDesktopW", (["in"], HWND, "hwnd"), (["out"], POINTER(POINTER(IVirtualDesktop)), "pDesktop"),),
            STDMETHOD(HRESULT, "MoveDesktop", (POINTER(IVirtualDesktop), HWND, INT)),
            COMMETHOD([], HRESULT, "RemoveDesktop", (["in"], POINTER(IVirtualDesktop), "destroyDesktop"), (["in"], POINTER(IVirtualDesktop), "fallbackDesktop")),
            COMMETHOD([], HRESULT, "FindDesktop", (["in"], POINTER(GUID), "pGuid"), (["out"], POINTER(POINTER(IVirtualDesktop)), "pDesktop")),
            STDMETHOD(HRESULT, "Unknown", (POINTER(IVirtualDesktop), POINTER(POINTER(IObjectArray)), POINTER(POINTER(IObjectArray)))),
            COMMETHOD([], HRESULT, "SetName", (["in"], POINTER(IVirtualDesktop), "pDesktop"), (["in"], HSTRING, "name")),
            COMMETHOD([], HRESULT, "SetWallpaper", (["in"], POINTER(IVirtualDesktop), "pDesktop"), (["in"], HSTRING, "path")),
            COMMETHOD([], HRESULT, "SetWallpaperForAllDesktops", (["in"], HSTRING, "path")),
            COMMETHOD([], HRESULT, "CopyDesktopState", (["in"], POINTER(IApplicationView), "pView0"), (["in"], POINTER(IApplicationView), "pView0")),
            COMMETHOD([], HRESULT, "GetDesktopPerMonitor", (["out"], POINTER(BOOL), "state")),
            COMMETHOD([], HRESULT, "SetDesktopPerMonitor", (["in"], BOOL, "state")),
        ]
    elif BUILD_OVER_21313:
        _methods_ = [
            COMMETHOD([], HRESULT, "GetCount", (["in"], HWND, "hwnd"), (["out"], POINTER(UINT), "pCount"),),
            STDMETHOD(HRESULT, "MoveViewToDesktop", (POINTER(IApplicationView), POINTER(IVirtualDesktop))),
            STDMETHOD(HRESULT, "CanViewMoveDesktops", (POINTER(IApplicationView), POINTER(UINT))),
            COMMETHOD([], HRESULT, "GetCurrentDesktop", (["in"], HWND, "hwnd"), (["out"], POINTER(POINTER(IVirtualDesktop)), "pDesktop"),),
            COMMETHOD([], HRESULT, "GetDesktops", (["in"], HWND, "hwnd"), (["out"], POINTER(POINTER(IObjectArray)), "array")),
            STDMETHOD(HRESULT, "GetAdjacentDesktop", (POINTER(IVirtualDesktop), AdjacentDesktop, POINTER(POINTER(IVirtualDesktop)),)),
            STDMETHOD(HRESULT, "SwitchDesktop", (HWND, POINTER(IVirtualDesktop),)),
            COMMETHOD([], HRESULT, "CreateDesktopW", (["in"], HWND, "hwnd"), (["out"], POINTER(POINTER(IVirtualDesktop)), "pDesktop"),),
            STDMETHOD(HRESULT, "MoveDesktop", (POINTER(IVirtualDesktop), HWND, INT)),
            COMMETHOD([], HRESULT, "RemoveDesktop", (["in"], POINTER(IVirtualDesktop), "destroyDesktop"), (["in"], POINTER(IVirtualDesktop), "fallbackDesktop")),
            COMMETHOD([], HRESULT, "FindDesktop", (["in"], POINTER(GUID), "pGuid"), (["out"], POINTER(POINTER(IVirtualDesktop)), "pDesktop")),
            STDMETHOD(HRESULT, "Unknown", (POINTER(IVirtualDesktop), POINTER(POINTER(IObjectArray)), POINTER(POINTER(IObjectArray)))),
            COMMETHOD([], HRESULT, "SetName", (["in"], POINTER(IVirtualDesktop), "pDesktop"), (["in"], HSTRING, "name")),
            COMMETHOD([], HRESULT, "SetWallpaper", (["in"], POINTER(IVirtualDesktop), "pDesktop"), (["in"], HSTRING, "path")),
            COMMETHOD([], HRESULT, "SetWallpaperForAllDesktops", (["in"], HSTRING, "path")),
            COMMETHOD([], HRESULT, "CopyDesktopState", (["in"], POINTER(IApplicationView), "pView0"), (["in"], POINTER(IApplicationView), "pView0")),
            COMMETHOD([], HRESULT, "GetDesktopPerMonitor", (["out"], POINTER(BOOL), "state")),
            COMMETHOD([], HRESULT, "SetDesktopPerMonitor", (["in"], BOOL, "state")),
        ]
    elif BUILD_OVER_20231:
        _methods_ = [
            COMMETHOD([], HRESULT, "GetCount", (["in"], HWND, "hwnd"), (["out"], POINTER(UINT), "pCount"),),
            STDMETHOD(HRESULT, "MoveViewToDesktop", (POINTER(IApplicationView), POINTER(IVirtualDesktop))),
            STDMETHOD(HRESULT, "CanViewMoveDesktops", (POINTER(IApplicationView), POINTER(UINT))),
            COMMETHOD([], HRESULT, "GetCurrentDesktop", (["in"], HWND, "hwnd"), (["out"], POINTER(POINTER(IVirtualDesktop)), "pDesktop"),),
            COMMETHOD([], HRESULT, "GetDesktops", (["in"], HWND, "hwnd"), (["out"], POINTER(POINTER(IObjectArray)), "array")),
            STDMETHOD(HRESULT, "GetAdjacentDesktop", (POINTER(IVirtualDesktop), AdjacentDesktop, POINTER(POINTER(IVirtualDesktop)),)),
            STDMETHOD(HRESULT, "SwitchDesktop", (HWND, POINTER(IVirtualDesktop),)),
            COMMETHOD([], HRESULT, "CreateDesktopW", (["in"], HWND, "hwnd"), (["out"], POINTER(POINTER(IVirtualDesktop)), "pDesktop"),),
            COMMETHOD([], HRESULT, "RemoveDesktop", (["in"], POINTER(IVirtualDesktop), "destroyDesktop"), (["in"], POINTER(IVirtualDesktop), "fallbackDesktop")),
            COMMETHOD([], HRESULT, "FindDesktop", (["in"], POINTER(GUID), "pGuid"), (["out"], POINTER(POINTER(IVirtualDesktop)), "pDesktop")),
        ]
    else:
        _methods_ = [
            COMMETHOD([], HRESULT, "GetCount", (["out"], POINTER(UINT), "pCount"),),
            STDMETHOD(HRESULT, "MoveViewToDesktop", (POINTER(IApplicationView), POINTER(IVirtualDesktop))),
            STDMETHOD(HRESULT, "CanViewMoveDesktops", (POINTER(IApplicationView), POINTER(UINT))),
            COMMETHOD([], HRESULT, "GetCurrentDesktop", (["out"], POINTER(POINTER(IVirtualDesktop)), "pDesktop"),),
            COMMETHOD([], HRESULT, "GetDesktops", (["out"], POINTER(POINTER(IObjectArray)), "array")),
            STDMETHOD(HRESULT, "GetAdjacentDesktop", (POINTER(IVirtualDesktop), AdjacentDesktop, POINTER(POINTER(IVirtualDesktop)),)),
            STDMETHOD(HRESULT, "SwitchDesktop", (POINTER(IVirtualDesktop),)),
            COMMETHOD([], HRESULT, "CreateDesktopW", (["out"], POINTER(POINTER(IVirtualDesktop)), "pDesktop"),),
            COMMETHOD([], HRESULT, "RemoveDesktop", (["in"], POINTER(IVirtualDesktop), "destroyDesktop"), (["in"], POINTER(IVirtualDesktop), "fallbackDesktop")),
            COMMETHOD([], HRESULT, "FindDesktop", (["in"], POINTER(GUID), "pGuid"), (["out"], POINTER(POINTER(IVirtualDesktop)), "pDesktop")),
        ]

    def get_all_desktops(self) -> IObjectArray:
        if BUILD_OVER_22621:
            return self.GetDesktops()
        elif BUILD_OVER_22449:
            # See https://github.com/mrob95/pyvda/issues/15#issuecomment-1146642608 for
            # discussion/theorising about this. Unclear at this point where MS are
            # going with it. Could probably be zero but don't want to break it when
            # I don't have a test machine.
            return self.GetDesktops(1)
        elif BUILD_OVER_20231:
            return self.GetDesktops(0)
        else:
            return self.GetDesktops()

    def get_current_desktop(self) -> IVirtualDesktop:
        if BUILD_OVER_22621:
            return self.GetCurrentDesktop()
        elif BUILD_OVER_20231:
            return self.GetCurrentDesktop(0)
        else:
            return self.GetCurrentDesktop()

    def create_desktop(self) -> IVirtualDesktop:
        if BUILD_OVER_22621:
            return self.CreateDesktopW()
        elif BUILD_OVER_20231:
            return self.CreateDesktopW(0)
        else:
            return self.CreateDesktopW()

    def switch_desktop(self, target: IVirtualDesktop) -> IVirtualDesktop:
        if BUILD_OVER_22621:
            return self.SwitchDesktop(target)
        elif BUILD_OVER_20231:
            return self.SwitchDesktop(0, target)
        else:
            return self.SwitchDesktop(target)


GUID_IVirtualDesktopManagerInternal2 = GUID("{0F3A72B0-4566-487E-9A33-4ED302F6D6CE}")
class IVirtualDesktopManagerInternal2(IUnknown):
    _iid_ = GUID_IVirtualDesktopManagerInternal2
    _methods_ = [
        COMMETHOD([], HRESULT, "GetCount", (["out"], POINTER(UINT), "pCount"),),
        STDMETHOD(HRESULT, "MoveViewToDesktop", (POINTER(IApplicationView), POINTER(IVirtualDesktop))),
        STDMETHOD(HRESULT, "CanViewMoveDesktops", (POINTER(IApplicationView), POINTER(UINT))),
        COMMETHOD([], HRESULT, "GetCurrentDesktop", (["out"], POINTER(POINTER(IVirtualDesktop)), "pDesktop"),),
        COMMETHOD([], HRESULT, "GetDesktops", (["out"], POINTER(POINTER(IObjectArray)), "array")),
        STDMETHOD(HRESULT, "GetAdjacentDesktop", (POINTER(IVirtualDesktop), AdjacentDesktop, POINTER(POINTER(IVirtualDesktop)),)),
        STDMETHOD(HRESULT, "SwitchDesktop", (POINTER(IVirtualDesktop),)),
        COMMETHOD([], HRESULT, "CreateDesktopW", (["out"], POINTER(POINTER(IVirtualDesktop)), "pDesktop"),),
        COMMETHOD([], HRESULT, "RemoveDesktop", (["in"], POINTER(IVirtualDesktop), "destroyDesktop"), (["in"], POINTER(IVirtualDesktop), "fallbackDesktop")),
        COMMETHOD([], HRESULT, "FindDesktop", (["in"], POINTER(GUID), "pGuid"), (["out"], POINTER(POINTER(IVirtualDesktop)), "pDesktop")),
        STDMETHOD(HRESULT, "Unknown", (POINTER(IVirtualDesktop), POINTER(POINTER(IObjectArray)), POINTER(POINTER(IObjectArray)))),
        COMMETHOD([], HRESULT, "SetName", (["in"], POINTER(IVirtualDesktop), "pDesktop"), (["in"], HSTRING, "name")),
    ]


# aa509086-5ca9-4c25-8f95-589d3c07b48a ?
# HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Interface\{A5CD92FF-29BE-454C-8D04-D82879FB3F1B}
class IVirtualDesktopManager(IUnknown):
    _iid_ = GUID("{A5CD92FF-29BE-454C-8D04-D82879FB3F1B}")
    _methods_ = [
        COMMETHOD([], HRESULT, "IsWindowOnCurrentVirtualDesktop",
            (["in"], HWND, "hwnd"),
            (["out"], POINTER(BOOL), "isOnCurrent"),
        ),
        STDMETHOD(HRESULT, "GetWindowDesktopId", (HWND, POINTER(GUID))),
        STDMETHOD(HRESULT, "MoveWindowToDesktop", (HWND, REFGUID)),
    ]


class IVirtualDesktopPinnedApps(IUnknown):
    _iid_ = GUID("{4CE81583-1E4C-4632-A621-07A53543148F}")
    _methods_ = [
        # IVirtualDesktopPinnedApps methods
        COMMETHOD([], HRESULT, "IsAppIdPinned",
            (["in"], LPCWSTR, "appId"),
            (["out"], POINTER(BOOL), "isPinned"),
        ),
        STDMETHOD(HRESULT, "PinAppID", (LPCWSTR,)),
        STDMETHOD(HRESULT, "UnpinAppID", (LPCWSTR,)),
        COMMETHOD([], HRESULT, "IsViewPinned",
            (["in"], POINTER(IApplicationView), "pView"),
            (["out"], POINTER(BOOL), "isPinned"),
        ),
        STDMETHOD(HRESULT, "PinView", (POINTER(IApplicationView),)),
        STDMETHOD(HRESULT, "UnpinView", (POINTER(IApplicationView),)),
    ]


# In registry: Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Interface\{1841C6D7-4F9D-42C0-AF41-8747538F10E5}
class IApplicationViewCollection(IUnknown):
    _iid_ = GUID("{1841C6D7-4F9D-42C0-AF41-8747538F10E5}")
    _methods_ = [
        # IApplicationViewCollection methods
        STDMETHOD(HRESULT, "GetViews", (POINTER(POINTER(IObjectArray)),)),
        # STDMETHOD(HRESULT, "GetViewsByZOrder", (POINTER(POINTER(IObjectArray)),)),
        COMMETHOD([], HRESULT, "GetViewsByZOrder", (["out"], POINTER(POINTER(IObjectArray)), "array")),
        STDMETHOD(HRESULT, "GetViewsByAppUserModelId", (LPCWSTR, POINTER(POINTER(IObjectArray)))),
        COMMETHOD([], HRESULT, "GetViewForHwnd",
            (["in"], HWND, "hwnd"),
            (["out"], POINTER(POINTER(IApplicationView)), "pView")
        ),
        STDMETHOD(HRESULT, "GetViewForApplication", (POINTER(IImmersiveApplication), POINTER(POINTER(IApplicationView)))),
        STDMETHOD(HRESULT, "GetViewForAppUserModelId", (LPCWSTR, POINTER(POINTER(IApplicationView)))),
        COMMETHOD([], HRESULT, "GetViewInFocus", (["out"], POINTER(POINTER(IApplicationView)), "view")),
        STDMETHOD(HRESULT, "Unknown1", (POINTER(POINTER(IApplicationView)),)),
        STDMETHOD(HRESULT, "RefreshCollection", ()),
        STDMETHOD(HRESULT, "RegisterForApplicationViewChanges", (POINTER(IApplicationViewChangeListener), POINTER(DWORD))),
        # Removed in 1809
        # STDMETHOD(HRESULT, "RegisterForApplicationViewPositionChanges", (POINTER(IApplicationViewChangeListener), POINTER(DWORD),)),
        STDMETHOD(HRESULT, "UnregisterForApplicationViewChanges", (DWORD,)),
    ]
