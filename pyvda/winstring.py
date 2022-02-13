# https://github.com/ninthDevilHAUNSTER/ArknightsAutoHelper/blob/fc2737c7cc4400f5a90763a9eef7c9a702c1f2be/rotypes/types.py#L17
#
# MIT License

# Copyright (c) 2019 ShaoBaoBaoEr

# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:

# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.

# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.
#

import ctypes
import weakref

E_FAIL = -2147467259  # 0x80004005L
E_NOTIMPL = -2147467263  # 0x80004001L
E_NOINTERFACE = -2147467262  # 0x80004002L
E_BOUNDS = -2147483637  # 0x8000000BL


def check_hresult(hr):
    # print('HRESULT = 0x%08X' % (hr & 0xFFFFFFFF))
    if (hr & 0x80000000) != 0:
        if hr == E_NOTIMPL:
            raise NotImplementedError
        elif hr == E_NOINTERFACE:
            raise TypeError("E_NOINTERFACE")
        elif hr == E_BOUNDS:
            raise IndexError  # for old style iterator protocol
        e = OSError("[HRESULT 0x%08X] %s" % (hr & 0xFFFFFFFF, ctypes.FormatError(hr)))
        e.winerror = hr & 0xFFFFFFFF
        raise e
    return hr


combase = ctypes.windll.LoadLibrary("combase.dll")
WindowsCreateString = combase.WindowsCreateString
WindowsCreateString.argtypes = (ctypes.c_void_p, ctypes.c_uint32, ctypes.POINTER(ctypes.c_void_p))
WindowsCreateString.restype = check_hresult

WindowsDeleteString = combase.WindowsDeleteString
WindowsDeleteString.argtypes = (ctypes.c_void_p,)
WindowsDeleteString.restype = check_hresult

WindowsGetStringRawBuffer = combase.WindowsGetStringRawBuffer
WindowsGetStringRawBuffer.argtypes = (ctypes.c_void_p, ctypes.POINTER(ctypes.c_uint32))
WindowsGetStringRawBuffer.restype = ctypes.c_void_p



class HSTRING(ctypes.c_void_p):
    def __init__(self, s=None):
        super().__init__()
        if s is None or len(s) == 0:
            self.value = None
            return
        u16str = s.encode("utf-16-le") + b"\x00\x00"
        u16len = (len(u16str) // 2) - 1
        WindowsCreateString(u16str, ctypes.c_uint32(u16len), ctypes.byref(self))
        self._finalizer = weakref.finalize(self, WindowsDeleteString, self.value)  # only register finalizer if we created the string

    def __str__(self):
        if self.value is None:
            return ""
        length = ctypes.c_uint32()
        ptr = WindowsGetStringRawBuffer(self, ctypes.byref(length))
        return ctypes.wstring_at(ptr, length.value)

    def __repr__(self):
        return "HSTRING(%s)" % repr(str(self))
