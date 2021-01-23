import win32gui
from pyvda import *

def test_count():
    count = GetDesktopCount()
    assert 0 < count
    assert count < 30

def test_current():
    current = GetCurrentDesktopNumber()
    assert 0 < current
    assert current < 30

current_window_handle = win32gui.GetForegroundWindow()
current_desktop_number = GetCurrentDesktopNumber()

def test_move():
    MoveWindowToDesktopNumber(current_window_handle, 1)

def test_go():
    GoToDesktopNumber(1)
    assert GetCurrentDesktopNumber() == 1

def test_get_number():
    assert GetWindowDesktopNumber(current_window_handle) == 1

def test_cleanup():
    MoveWindowToDesktopNumber(current_window_handle, current_desktop_number)
    GoToDesktopNumber(current_desktop_number)

def test_pin_app():
    assert not IsPinnedApp(current_window_handle)
    PinApp(current_window_handle)
    assert IsPinnedApp(current_window_handle)
    UnPinApp(current_window_handle)
    assert not IsPinnedApp(current_window_handle)

def test_pin_window():
    assert not IsPinnedWindow(current_window_handle)
    PinWindow(current_window_handle)
    assert IsPinnedWindow(current_window_handle)
    UnPinWindow(current_window_handle)
    assert not IsPinnedWindow(current_window_handle)