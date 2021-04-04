# py-VirtualDesktopAccessor

Implements a subset of the functionality offered by https://github.com/Ciantic/VirtualDesktopAccessor, written in pure Python and installable via pip.

## Installation
```
pip install pyvda
```

## Example usage
```python
from pyvda import AppView, get_apps_by_z_order, VirtualDesktop, get_virtual_desktops

number_of_active_desktops = len(get_virtual_desktops())
print(f"There are {number_of_active_desktops} active desktops")

current_desktop = VirtualDesktop.current()
print(f"Current desktop is number {current_desktop}")

current_window = AppView.current()
target_desktop = VirtualDesktop(5)
current_window.move(target_desktop)
print(f"Moved window {current_window.hwnd} to {target_desktop.number}")

print("Going to desktop number 5")
VirtualDesktop(5).go()

print("Pinning the current window")
AppView.current().pin()
```

## Documentation
Full API documentation can be found at

## Tips

Sometimes, after calling `GoToDesktopNumber` the focus will remain on the window in the previous desktop. This is at least partially fixed by calling:
```
from ctypes import windll
ASFW_ANY = -1
windll.user32.AllowSetForegroundWindow(ASFW_ANY)
```

before any call to `GoToDesktopNumber`.