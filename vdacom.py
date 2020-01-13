from ctypes import *
from ctypes.wintypes import *
from comtypes import IUnknown, GUID, IID, COMMETHOD, CoInitialize, CoCreateInstance, cast, CLSCTX_LOCAL_SERVER

from comtypes.typeinfo import ITypeInfo

REFGUID = POINTER(GUID)
REFIID = POINTER(GUID)

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
        COMMETHOD([], HRESULT, "GetCount",
            (["out"], POINTER(UINT), "pcObjects"),
        ),
        COMMETHOD([], HRESULT, "GetAt",
            (["in"],  UINT, "uiIndex"),
            (["in"],  REFIID, "riid"),
            (["out"], POINTER(LPVOID), "ppv"),
        ),
    ]

# In registry: Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Interface\{FF72FFDD-BE7E-43FC-9C03-AD81681E88E4}
class IVirtualDesktop(IUnknown):
    _iid_ = GUID("{FF72FFDD-BE7E-43FC-9C03-AD81681E88E4}")
    _methods_ = [
	virtual HRESULT STDMETHODCALLTYPE IsViewVisible(
		IApplicationView *pView,
		int *pfVisible) = 0;

	virtual HRESULT STDMETHODCALLTYPE GetID(
		GUID *pGuid) = 0;
    ]


# HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Interface\{F31574D6-B682-4CDC-BD56-1827860ABEC6}
class IVirtualDesktopManagerInternal(IUnknown):
    _iid_ = GUID("{F31574D6-B682-4CDC-BD56-1827860ABEC6}")
    _methods_ = [
	virtual HRESULT STDMETHODCALLTYPE GetCount(
		UINT *pCount) = 0;

	virtual HRESULT STDMETHODCALLTYPE MoveViewToDesktop(
		IApplicationView *pView,
		IVirtualDesktop *pDesktop) = 0;

	# Since build 10240
	virtual HRESULT STDMETHODCALLTYPE CanViewMoveDesktops(
		IApplicationView *pView,
		int *pfCanViewMoveDesktops) = 0;

	virtual HRESULT STDMETHODCALLTYPE GetCurrentDesktop(
		IVirtualDesktop** desktop) = 0;

	virtual HRESULT STDMETHODCALLTYPE GetDesktops(
		IObjectArray **ppDesktops) = 0;

	virtual HRESULT STDMETHODCALLTYPE GetAdjacentDesktop(
		IVirtualDesktop *pDesktopReference,
		AdjacentDesktop uDirection,
		IVirtualDesktop **ppAdjacentDesktop) = 0;

	virtual HRESULT STDMETHODCALLTYPE SwitchDesktop(
		IVirtualDesktop *pDesktop) = 0;

	virtual HRESULT STDMETHODCALLTYPE CreateDesktopW(
		IVirtualDesktop **ppNewDesktop) = 0;

	virtual HRESULT STDMETHODCALLTYPE RemoveDesktop(
		IVirtualDesktop *pRemove,
		IVirtualDesktop *pFallbackDesktop) = 0;

	# Since build 10240
	virtual HRESULT STDMETHODCALLTYPE FindDesktop(
		GUID *desktopId,
		IVirtualDesktop **ppDesktop) = 0;
    ]

# aa509086-5ca9-4c25-8f95-589d3c07b48a ?
# HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Interface\{A5CD92FF-29BE-454C-8D04-D82879FB3F1B}
class IVirtualDesktopManager(IUnknown):
    _iid_ = GUID("{A5CD92FF-29BE-454C-8D04-D82879FB3F1B}")
    _methods_ = [
	virtual HRESULT STDMETHODCALLTYPE IsWindowOnCurrentVirtualDesktop(
		/* [in] */ __RPC__in HWND topLevelWindow,
		/* [out] */ __RPC__out BOOL *onCurrentDesktop) = 0;

	virtual HRESULT STDMETHODCALLTYPE GetWindowDesktopId(
		/* [in] */ __RPC__in HWND topLevelWindow,
		/* [out] */ __RPC__out GUID *desktopId) = 0;

	virtual HRESULT STDMETHODCALLTYPE MoveWindowToDesktop(
		/* [in] */ __RPC__in HWND topLevelWindow,
		/* [in] */ __RPC__in REFGUID desktopId) = 0;
    ]

CoInitialize()