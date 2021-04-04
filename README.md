# Python Virtual Desktop Accessor
Python module providing programmatic access to most of the settings accessed through the [Windows 10 task view](https://en.wikipedia.org/wiki/Task_View).
Including switching virtual desktops, moving windows between virtual desktops, pinning windows and listing the windows on a desktop.

Originally based on https://github.com/Ciantic/VirtualDesktopAccessor.

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
Full API documentation can be found at [Read the Docs](https://pyvda.readthedocs.io/en/latest/index.html)
