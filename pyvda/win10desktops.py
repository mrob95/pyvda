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
from comtypes import (
    IUnknown,
    GUID,
    STDMETHOD,
    COMMETHOD,
)

# See https://github.com/Ciantic/VirtualDesktopAccessor/issues/33
# https://github.com/mzomparelli/zVirtualDesktop/wiki
WINDOWS_BUILD = sys.getwindowsversion().build
BUILD_OVER_20231 = WINDOWS_BUILD >= 20231


CLSID_ImmersiveShell = GUID("{C2F03A33-21F5-47FA-B4BB-156362A2F239}")
CLSID_VirtualDesktopManagerInternal = GUID("{C5E0CDCA-7B6E-41B2-9FC4-D93975CC467B}")
CLSID_IVirtualDesktopManager = GUID("{AA509086-5CA9-4C25-8F95-589D3C07B48A}")
CLSID_VirtualDesktopPinnedApps = GUID("{B5A399E7-1C87-46B8-88E9-FC5747B171BD}")

# Ignore following API's:
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


class HSTRING(Structure):
    _fields_ = [("value", WCHAR), ("size", UINT)]


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
    # STDMETHOD(HRESULT, "GetVisibility", (POINTER(UINT),)),
    COMMETHOD([], HRESULT, "GetVisibility", (["out"], POINTER(UINT), "pVisible")),
    STDMETHOD(HRESULT, "SetCloak", (APPLICATION_VIEW_CLOAK_TYPE, UINT,)),
    STDMETHOD(HRESULT, "GetPosition", (REFIID, POINTER(LPVOID),)),
    STDMETHOD(HRESULT, "SetPosition", (POINTER(IApplicationViewPosition),)),
    STDMETHOD(HRESULT, "InsertAfterWindow", (HWND,)),
    STDMETHOD(HRESULT, "GetExtendedFramePosition", (POINTER(RECT),)),
    # STDMETHOD(HRESULT, "GetAppUserModelId", (POINTER(PWSTR),)),
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
    STDMETHOD(HRESULT, "Unknown1", (POINTER(UINT),)),
    STDMETHOD(HRESULT, "Unknown2", (POINTER(UINT),)),
    STDMETHOD(HRESULT, "Unknown3", (POINTER(UINT),)),
    STDMETHOD(HRESULT, "Unknown4", (UINT,)),
    STDMETHOD(HRESULT, "Unknown5", (POINTER(UINT),)),
    STDMETHOD(HRESULT, "Unknown6", (UINT,)),
    STDMETHOD(HRESULT, "Unknown7", ()),
    STDMETHOD(HRESULT, "Unknown8", (POINTER(UINT),)),
    STDMETHOD(HRESULT, "Unknown9", (UINT,)),
    STDMETHOD(HRESULT, "Unknown10", (UINT, UINT,)),
    STDMETHOD(HRESULT, "Unknown11", (UINT,)),
    STDMETHOD(HRESULT, "Unknown12", (POINTER(SIZE),)),
]

if BUILD_OVER_20231:
    GUID_IVirtualDesktop = GUID("{62FDF88B-11CA-4AFB-8BD8-2296DFAE49E2}")
else:
    GUID_IVirtualDesktop = GUID("{FF72FFDD-BE7E-43FC-9C03-AD81681E88E4}")

# In registry: Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Interface\{FF72FFDD-BE7E-43FC-9C03-AD81681E88E4}
class IVirtualDesktop(IUnknown):
    _iid_ = GUID_IVirtualDesktop
    _methods_ = [
        STDMETHOD(HRESULT, "IsViewVisible", (POINTER(IApplicationView), POINTER(UINT))),
        COMMETHOD([], HRESULT, "GetID", (["out"], POINTER(GUID), "pGuid"),),
    ]


if BUILD_OVER_20231:
    GUID_IVirtualDesktopManagerInternal = GUID("{094AFE11-44F2-4BA0-976F-29A97E263EE0}")
else:
    GUID_IVirtualDesktopManagerInternal = GUID("{F31574D6-B682-4CDC-BD56-1827860ABEC6}")

# HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Interface\{F31574D6-B682-4CDC-BD56-1827860ABEC6}
class IVirtualDesktopManagerInternal(IUnknown):
    _iid_ = GUID_IVirtualDesktopManagerInternal
    _methods_ = [
        COMMETHOD([], HRESULT, "GetCount", (["out"], POINTER(UINT), "pCount"),),
        STDMETHOD(HRESULT, "MoveViewToDesktop", (POINTER(IApplicationView), POINTER(IVirtualDesktop))),
        # Since build 10240
        STDMETHOD(HRESULT, "CanViewMoveDesktops", (POINTER(IApplicationView), POINTER(UINT))),
        STDMETHOD(HRESULT, "GetCurrentDesktop", (POINTER(POINTER(IVirtualDesktop)),)),
        COMMETHOD([], HRESULT, "GetDesktops", (["out"], POINTER(POINTER(IObjectArray)), "array")),
        STDMETHOD(HRESULT, "GetAdjacentDesktop", (
            POINTER(IVirtualDesktop), AdjacentDesktop, POINTER(POINTER(IVirtualDesktop)),
        )),
        STDMETHOD(HRESULT, "SwitchDesktop", (POINTER(IVirtualDesktop),)),
        STDMETHOD(HRESULT, "CreateDesktopW", (POINTER(POINTER(IVirtualDesktop)),)),
        STDMETHOD(HRESULT, "RemoveDesktop", (POINTER(IVirtualDesktop), POINTER(IVirtualDesktop))),
        # Since build 10240
        STDMETHOD(HRESULT, "FindDesktop", (POINTER(GUID), POINTER(POINTER(IVirtualDesktop)))),
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
        # STDMETHOD(HRESULT, "IsWindowOnCurrentVirtualDesktop", (HWND, POINTER(BOOL))),
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
        STDMETHOD(HRESULT, "GetViewInFocus", (POINTER(POINTER(IApplicationView)),)),
        STDMETHOD(HRESULT, "Unknown1", (POINTER(POINTER(IApplicationView)),)),
        STDMETHOD(HRESULT, "RefreshCollection", ()),
        STDMETHOD(HRESULT, "RegisterForApplicationViewChanges", (POINTER(IApplicationViewChangeListener), POINTER(DWORD))),
        # Removed in 1809
        # STDMETHOD(HRESULT, "RegisterForApplicationViewPositionChanges", (POINTER(IApplicationViewChangeListener), POINTER(DWORD),)),
        STDMETHOD(HRESULT, "UnregisterForApplicationViewChanges", (DWORD,)),
    ]

