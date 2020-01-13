from ctypes import *
from ctypes.wintypes import *
from comtypes import (
    IUnknown,
    GUID,
    IID,
    COMMETHOD,
    CoInitialize,
    CoCreateInstance,
    cast,
    CLSCTX_LOCAL_SERVER,
)

REFGUID = POINTER(GUID)
REFIID = POINTER(GUID)
IApplicationView = UINT
AdjacentDesktop = UINT

CLSID_ImmersiveShell = GUID("{C2F03A33-21F5-47FA-B4BB-156362A2F239}")
CLSID_VirtualDesktopManagerInternal = GUID("{C5E0CDCA-7B6E-41B2-9FC4-D93975CC467B}")


class IServiceProvider(IUnknown):
    _iid_ = GUID("{6D5140C1-7436-11CE-8034-00AA006009FA}")
    _methods_ = [
        COMMETHOD(
            [],
            HRESULT,
            "QueryService",
            (["in"], REFGUID, "guidService"),
            (["in"], REFIID, "riid"),
            (["out"], POINTER(LPVOID), "ppvObject"),
        ),
    ]


class IObjectArray(IUnknown):
    _iid_ = GUID("{92CA9DCD-5622-4BBA-A805-5E9F541BD8C9}")
    _methods_ = [
        COMMETHOD([], HRESULT, "GetCount", (["out"], POINTER(UINT), "pcObjects"),),
        COMMETHOD(
            [],
            HRESULT,
            "GetAt",
            (["in"], UINT, "uiIndex"),
            (["in"], REFIID, "riid"),
            (["out"], POINTER(LPVOID), "ppv"),
        ),
    ]


# In registry: Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Interface\{FF72FFDD-BE7E-43FC-9C03-AD81681E88E4}
class IVirtualDesktop(IUnknown):
    _iid_ = GUID("{FF72FFDD-BE7E-43FC-9C03-AD81681E88E4}")
    _methods_ = [
        COMMETHOD(
            [],
            HRESULT,
            "IsViewVisible",
            (["in"], POINTER(IApplicationView), "pView"),
            (["out"], POINTER(UINT), "pfVisible"),
        ),
        COMMETHOD([], HRESULT, "GetID", (["in"], POINTER(GUID), "pGuid"),),
    ]


# HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Interface\{F31574D6-B682-4CDC-BD56-1827860ABEC6}
class IVirtualDesktopManagerInternal(IUnknown):
    _iid_ = GUID("{F31574D6-B682-4CDC-BD56-1827860ABEC6}")
    _methods_ = [
        COMMETHOD([], HRESULT, "GetCount", (["in"], POINTER(UINT), "pCount"),),
        COMMETHOD(
            [],
            HRESULT,
            "MoveViewToDesktop",
            (["in"], POINTER(IApplicationView), "pView"),
            (["in"], POINTER(IVirtualDesktop), "pDesktop"),
        ),
        # Since build 10240
        COMMETHOD(
            [],
            HRESULT,
            "CanViewMoveDesktops",
            (["in"], POINTER(IApplicationView), "pView"),
            (["in"], POINTER(UINT), "pfCanViewMoveDesktops"),
        ),
        COMMETHOD(
            [],
            HRESULT,
            "GetCurrentDesktop",
            (["in"], POINTER(POINTER(IVirtualDesktop)), "desktop"),
        ),
        COMMETHOD(
            [],
            HRESULT,
            "GetDesktops",
            (["in"], POINTER(POINTER(IObjectArray)), "ppDesktops"),
        ),
        COMMETHOD(
            [],
            HRESULT,
            "GetAdjacentDesktop",
            (["in"], POINTER(IVirtualDesktop), "pDesktopReference"),
            (["in"], AdjacentDesktop, "uDirection"),
            (["out"], POINTER(POINTER(IVirtualDesktop)), "ppAdjacentDesktop"),
        ),
        COMMETHOD(
            [],
            HRESULT,
            "SwitchDesktop",
            (["in"], POINTER(IVirtualDesktop), "pDesktop"),
        ),
        COMMETHOD(
            [],
            HRESULT,
            "CreateDesktopW",
            (["in"], POINTER(POINTER(IVirtualDesktop)), "ppNewDesktop"),
        ),
        COMMETHOD(
            [],
            HRESULT,
            "RemoveDesktop",
            (["in"], POINTER(IVirtualDesktop), "pRemove"),
            (["in"], POINTER(IVirtualDesktop), "pFallbackDesktop"),
        ),
        # Since build 10240
        COMMETHOD(
            [],
            HRESULT,
            "FindDesktop",
            (["in"], POINTER(GUID), "desktopId"),
            (["out"], POINTER(POINTER(IVirtualDesktop)), "ppDesktop"),
        ),
    ]


# aa509086-5ca9-4c25-8f95-589d3c07b48a ?
# HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Interface\{A5CD92FF-29BE-454C-8D04-D82879FB3F1B}
class IVirtualDesktopManager(IUnknown):
    _iid_ = GUID("{A5CD92FF-29BE-454C-8D04-D82879FB3F1B}")
    _methods_ = [
        COMMETHOD(
            [],
            HRESULT,
            "IsWindowOnCurrentVirtualDesktop",
            (["in"], HWND, "topLevelWindow"),
            (["out"], POINTER(BOOL), "onCurrentDesktop"),
        ),
        COMMETHOD(
            [],
            HRESULT,
            "GetWindowDesktopId",
            (["in"], HWND, "topLevelWindow"),
            (["out"], POINTER(GUID), "desktopId"),
        ),
        COMMETHOD(
            [],
            HRESULT,
            "MoveWindowToDesktop",
            (["in"], HWND, "topLevelWindow"),
            (["in"], REFGUID, "desktopId"),
        ),
    ]

def _register_service():
    CoInitialize()
    pServiceProvider = CoCreateInstance(
        CLSID_ImmersiveShell, IServiceProvider, CLSCTX_LOCAL_SERVER
    )
    pManagerInternal = cast(
        pServiceProvider.QueryService(
            CLSID_VirtualDesktopManagerInternal, IVirtualDesktopManagerInternal._iid_,
        ),
        POINTER(IVirtualDesktopManagerInternal),
    )
    return pServiceProvider, pManagerInternal

def get_desktop_count():
    pServiceProvider, pManagerInternal = _register_service()
    array = POINTER(IObjectArray)()
    pManagerInternal.GetDesktops(array)
    return array.GetCount()

# IVirtualDesktop* _GetDesktopByNumber(int number)
def _get_desktop_by_number(number):
    pServiceProvider, pManagerInternal = _register_service()
    array = POINTER(IObjectArray)()
    found = POINTER(IVirtualDesktop)()
    pManagerInternal.GetDesktops(array)
    array.GetAt(number, IVirtualDesktop._iid_, found)
    return found

def _get_current_desktop():
    pServiceProvider, pManagerInternal = _register_service()
    found = POINTER(IVirtualDesktop)()
    pManagerInternal.GetCurrentDesktop(found)
    return found


def current_desktop_id():
    ...


def move_window_to(hwnd, number):
    ...


def go_to_desktop(number):
    ...
