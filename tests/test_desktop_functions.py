from pyvda import AppView, VirtualDesktop, get_virtual_desktops, get_apps_by_z_order
import time
import win32gui

current_window = AppView.current()
current_desktop = VirtualDesktop.current()

def test_move():
    current_window.move(VirtualDesktop(1))

def test_go():
    VirtualDesktop(1).go()
    assert VirtualDesktop.current().number == 1

def test_get_number():
    assert current_window.desktop.number == 1

def test_cleanup():
    current_window.move(current_desktop)
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

def test_z_order():
    apps = get_apps_by_z_order()
    assert apps[0] == AppView.current()

    assert len(get_apps_by_z_order(False, False)) > len(apps)

def test_switch_focus():
    apps = get_apps_by_z_order()
    if len(apps) == 1:
        raise Exception("For testing purposes, open another window!")

    ts = apps[0].get_activation_timestamp()

    apps[1].set_focus()
    time.sleep(1)
    assert AppView.current() == apps[1]
    apps[0].switch_to()
    time.sleep(1)
    assert AppView.current() == apps[0]

    assert apps[0].get_activation_timestamp() > ts

def test_visibility():
    cur = AppView.current()
    assert cur.is_shown_in_switchers()
    assert cur.is_visible()

def test_current():
    count = len(get_virtual_desktops())
    current = VirtualDesktop.current().number
    assert 0 < current < count

    hwnd = win32gui.GetForegroundWindow()
    assert AppView(hwnd) == AppView.current()
    assert AppView(hwnd).is_on_current_desktop()