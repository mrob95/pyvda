from ctypes import HRESULT, POINTER, c_ulonglong
from ctypes.wintypes import LPVOID, UINT, WCHAR
from typing import Any, Iterator

from comtypes import COMMETHOD, GUID, STDMETHOD, IUnknown

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
        self.GetAt(i, cls._iid_, item) # type: ignore
        return item

    def iter(self, cls: Any) -> Iterator[Any]:
        for i in range(self.GetCount()): # type: ignore
            yield self.get_at(i, cls)
