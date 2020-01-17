# py-VirtualDesktopAccessor

Implements a subset of the functionality offered by https://github.com/Ciantic/VirtualDesktopAccessor, written in pure Python and installable via pip.

```
pip install pyvda
```

## Usage

The following functions are implemented. The only deliberate difference with the behaviour of Ciantic's original DLL is that desktops here are 1-indexed, as this reflects the numbers shown in the task view.

```python
def GetCurrentDesktopNumber() -> int:
def GetDesktopCount() -> int:
def MoveWindowToDesktopNumber(hwnd: int, number: int) -> None:
def GoToDesktopNumber(number: int) -> None:
def GetWindowDesktopNumber(hwnd: int) -> int:
```

Example usage:
```python
import pyvda
import win32gui

number_of_active_desktops = pyvda.GetDesktopCount()
current_desktop = pyvda.GetCurrentDesktopNumber()

current_window_handle = win32gui.GetForegroundWindow()
pyvda.MoveWindowToDesktopNumber(current_window_handle, 1)

pyvda.GoToDesktopNumber(3)

window_moved_to = pyvda.GetWindowDesktopNumber(current_window_handle)
```

## Tips

Sometimes, after calling `GoToDesktopNumber` the focus will remain on the window in the previous desktop. This is at least partially fixed by calling:
```
from ctypes import windll
ASFW_ANY = -1
windll.user32.AllowSetForegroundWindow(ASFW_ANY)
```

before any call to `GoToDesktopNumber`. More details [here](https://github.com/Ciantic/VirtualDesktopAccessor/issues/4) and [here](https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-allowsetforegroundwindow).
