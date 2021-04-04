import win32gui
from pyvda import AppView, VirtualDesktop, get_virtual_desktops

def test_count():
    count = len(get_virtual_desktops())
    assert 0 < count
    assert count < 30

def test_current():
    current = VirtualDesktop.current().number
    assert 0 < current
    assert current < 30

current_window = AppView.current()
current_desktop = VirtualDesktop.current()

def test_move():
    current_window.move_to_desktop(VirtualDesktop(1))

def test_go():
    VirtualDesktop(1).go()
    assert VirtualDesktop.current().number == 1

def test_get_number():
    assert current_window.desktop.number == 1

def test_cleanup():
    current_window.move_to_desktop(current_desktop)
    current_desktop.go()

def test_pin_app():
    assert not current_window.is_app_pinned()
    current_window.pin_app()
    assert current_window.is_app_pinned()
    current_window.unpin_app()
    assert not current_window.is_app_pinned()

def test_pin_window():
    assert not current_window.is_pinned()
    current_window.pin()
    assert current_window.is_pinned()
    current_window.unpin()
    assert not current_window.is_pinned()
