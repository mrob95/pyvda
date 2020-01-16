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
    IID,
    STDMETHOD,
    COMMETHOD,
    CoCreateInstance,
    CLSCTX_LOCAL_SERVER,
)

CLSID_ImmersiveShell = GUID("{C2F03A33-21F5-47FA-B4BB-156362A2F239}")
CLSID_VirtualDesktopManagerInternal = GUID("{C5E0CDCA-7B6E-41B2-9FC4-D93975CC467B}")
CLSID_IVirtualDesktopManager = GUID("{AA509086-5CA9-4C25-8F95-589D3C07B48A}")

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
    STDMETHOD(HRESULT, "GetThumbnailWindow", (POINTER(HWND),)),
    STDMETHOD(HRESULT, "GetMonitor", (POINTER(POINTER(IImmersiveMonitor)),)),
    STDMETHOD(HRESULT, "GetVisibility", (POINTER(UINT),)),
    STDMETHOD(HRESULT, "SetCloak", (APPLICATION_VIEW_CLOAK_TYPE, UINT,)),
    STDMETHOD(HRESULT, "GetPosition", (REFIID, POINTER(LPVOID),)),
    STDMETHOD(HRESULT, "SetPosition", (POINTER(IApplicationViewPosition),)),
    STDMETHOD(HRESULT, "InsertAfterWindow", (HWND,)),
    STDMETHOD(HRESULT, "GetExtendedFramePosition", (POINTER(RECT),)),
    STDMETHOD(HRESULT, "GetAppUserModelId", (POINTER(PWSTR),)),
    STDMETHOD(HRESULT, "SetAppUserModelId", (LPCWSTR,)),
    STDMETHOD(HRESULT, "IsEqualByAppUserModelId", (LPCWSTR, POINTER(UINT),)),
    STDMETHOD(HRESULT, "GetViewState", (POINTER(UINT),)),
    STDMETHOD(HRESULT, "SetViewState", (UINT,)),
    STDMETHOD(HRESULT, "GetNeediness", (POINTER(UINT),)),
    STDMETHOD(HRESULT, "GetLastActivationTimestamp", (POINTER(ULONGLONG),)),
    STDMETHOD(HRESULT, "SetLastActivationTimestamp", (ULONGLONG,)),
    COMMETHOD([], HRESULT, "GetVirtualDesktopId", (["out"], POINTER(GUID), "pGuid")),
    STDMETHOD(HRESULT, "SetVirtualDesktopId", (REFGUID,)),
    STDMETHOD(HRESULT, "GetShowInSwitchers", (POINTER(UINT),)),
    STDMETHOD(HRESULT, "SetShowInSwitchers", (UINT,)),
    STDMETHOD(HRESULT, "GetScaleFactor", (POINTER(UINT),)),
    STDMETHOD(HRESULT, "CanReceiveInput", (POINTER(BOOL),)),
    STDMETHOD(
        HRESULT,
        "GetCompatibilityPolicyType",
        (POINTER(APPLICATION_VIEW_COMPATIBILITY_POLICY),),
    ),
    STDMETHOD(
        HRESULT, "SetCompatibilityPolicyType", (APPLICATION_VIEW_COMPATIBILITY_POLICY,)
    ),
    # STDMETHOD(HRESULT, "GetPositionPriority", (POINTER(POINTER(IShellPositionerPriority)),)),
    # STDMETHOD(HRESULT, "SetPositionPriority", (POINTER(IShellPositionerPriority),)),
    STDMETHOD(
        HRESULT,
        "GetSizeConstraints",
        (POINTER(IImmersiveMonitor), POINTER(SIZE), POINTER(SIZE),),
    ),
    STDMETHOD(
        HRESULT, "GetSizeConstraintsForDpi", (UINT, POINTER(SIZE), POINTER(SIZE),)
    ),
    STDMETHOD(
        HRESULT,
        "SetSizeConstraintsForDpi",
        (POINTER(UINT), POINTER(SIZE), POINTER(SIZE),),
    ),
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


# In registry: Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Interface\{FF72FFDD-BE7E-43FC-9C03-AD81681E88E4}
class IVirtualDesktop(IUnknown):
    _iid_ = GUID("{FF72FFDD-BE7E-43FC-9C03-AD81681E88E4}")
    _methods_ = [
        STDMETHOD(HRESULT, "IsViewVisible", (POINTER(IApplicationView), POINTER(UINT))),
        COMMETHOD([], HRESULT, "GetID", (["out"], POINTER(GUID), "pGuid"),),
    ]


# HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Interface\{F31574D6-B682-4CDC-BD56-1827860ABEC6}
class IVirtualDesktopManagerInternal(IUnknown):
    _iid_ = GUID("{F31574D6-B682-4CDC-BD56-1827860ABEC6}")
    _methods_ = [
        COMMETHOD([], HRESULT, "GetCount", (["out"], POINTER(UINT), "pCount"),),
        STDMETHOD(
            HRESULT,
            "MoveViewToDesktop",
            (POINTER(IApplicationView), POINTER(IVirtualDesktop),),
        ),
        # Since build 10240
        STDMETHOD(
            HRESULT, "CanViewMoveDesktops", (POINTER(IApplicationView), POINTER(UINT),)
        ),
        STDMETHOD(HRESULT, "GetCurrentDesktop", (POINTER(POINTER(IVirtualDesktop)),)),
        STDMETHOD(HRESULT, "GetDesktops", (POINTER(POINTER(IObjectArray)),)),
        STDMETHOD(
            HRESULT,
            "GetAdjacentDesktop",
            (
                POINTER(IVirtualDesktop),
                AdjacentDesktop,
                POINTER(POINTER(IVirtualDesktop)),
            ),
        ),
        STDMETHOD(HRESULT, "SwitchDesktop", (POINTER(IVirtualDesktop),)),
        STDMETHOD(HRESULT, "CreateDesktopW", (POINTER(POINTER(IVirtualDesktop)),)),
        STDMETHOD(
            HRESULT,
            "RemoveDesktop",
            (POINTER(IVirtualDesktop), POINTER(IVirtualDesktop),),
        ),
        # Since build 10240
        STDMETHOD(
            HRESULT, "FindDesktop", (POINTER(GUID), POINTER(POINTER(IVirtualDesktop)))
        ),
    ]


# aa509086-5ca9-4c25-8f95-589d3c07b48a ?
# HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Interface\{A5CD92FF-29BE-454C-8D04-D82879FB3F1B}
class IVirtualDesktopManager(IUnknown):
    _iid_ = GUID("{A5CD92FF-29BE-454C-8D04-D82879FB3F1B}")
    _methods_ = [
        STDMETHOD(HRESULT, "IsWindowOnCurrentVirtualDesktop", (HWND, POINTER(BOOL))),
        STDMETHOD(HRESULT, "GetWindowDesktopId", (HWND, POINTER(GUID))),
        STDMETHOD(HRESULT, "MoveWindowToDesktop", (HWND, REFGUID)),
    ]


class IVirtualDesktopPinnedApps(IUnknown):
    _iid_ = GUID("{4CE81583-1E4C-4632-A621-07A53543148F}")
    _methods_ = [
        # IVirtualDesktopPinnedApps methods
        STDMETHOD(HRESULT, "IsAppIdPinned", (LPCWSTR, POINTER(BOOL))),
        STDMETHOD(HRESULT, "PinAppID", (LPCWSTR,)),
        STDMETHOD(HRESULT, "UnpinAppID", (LPCWSTR,)),
        STDMETHOD(HRESULT, "IsViewPinned", (POINTER(IApplicationView), POINTER(BOOL))),
        STDMETHOD(HRESULT, "PinView", (POINTER(IApplicationView),)),
        STDMETHOD(HRESULT, "UnpinView", (POINTER(IApplicationView),)),
    ]


# In registry: Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Interface\{1841C6D7-4F9D-42C0-AF41-8747538F10E5}
class IApplicationViewCollection(IUnknown):
    _iid_ = GUID("{1841C6D7-4F9D-42C0-AF41-8747538F10E5}")
    _methods_ = [
        # IApplicationViewCollection methods
        STDMETHOD(HRESULT, "GetViews", (POINTER(POINTER(IObjectArray)),)),
        STDMETHOD(HRESULT, "GetViewsByZOrder", (POINTER(POINTER(IObjectArray)),)),
        STDMETHOD(
            HRESULT,
            "GetViewsByAppUserModelId",
            (LPCWSTR, POINTER(POINTER(IObjectArray)),),
        ),
        STDMETHOD(
            HRESULT, "GetViewForHwnd", (HWND, POINTER(POINTER(IApplicationView)))
        ),
        STDMETHOD(
            HRESULT,
            "GetViewForApplication",
            (POINTER(IImmersiveApplication), POINTER(POINTER(IApplicationView))),
        ),
        STDMETHOD(
            HRESULT,
            "GetViewForAppUserModelId",
            (LPCWSTR, POINTER(POINTER(IApplicationView))),
        ),
        STDMETHOD(HRESULT, "GetViewInFocus", (POINTER(POINTER(IApplicationView)),)),
        STDMETHOD(HRESULT, "Unknown1", (POINTER(POINTER(IApplicationView)),)),
        STDMETHOD(HRESULT, "RefreshCollection", ()),
        STDMETHOD(
            HRESULT,
            "RegisterForApplicationViewChanges",
            (POINTER(IApplicationViewChangeListener), POINTER(DWORD),),
        ),
        # Removed in 1809
        # STDMETHOD(HRESULT, "RegisterForApplicationViewPositionChanges", (POINTER(IApplicationViewChangeListener), POINTER(DWORD),)),
        STDMETHOD(HRESULT, "UnregisterForApplicationViewChanges", (DWORD,)),
    ]

