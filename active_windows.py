'''
Functions for controlling windows on Windows OS.
'''
import win32gui
import win32con


def get_all_windows():
    '''
    Gets a list of all window titles.

    Returns
    ----------
    List
        List of window titles.
    '''
    window_hwnds = _get_windows()
    window_titles = []
    for hwnd in window_hwnds:
        title = win32gui.GetWindowText(hwnd)
        window_titles.append(title)
    return window_titles


def focus_window(window_title):
    '''
    Sets focus to window.

    Parameters
    ----------
    window_title : String
	    Window title.

    Returns
    ----------
    Boolean
	    True for success, False otherwise.
    '''
    hwnd = _get_window_hwnd(window_title)
    if hwnd is not None:
        win32gui.SetForegroundWindow(hwnd)
        win32gui.SetFocus(hwnd)
        return True
    else:
        return False


def get_focused_window_title():
    '''
    Returns currently focused Window title.

    Parameters
    ----------
    window_title : String
	    Window title.

    Returns
    ----------
    String
	    Focused window name. Returns None if no windows are active.
    '''
    hwnd = win32gui.GetForegroundWindow()
    return win32gui.GetWindowText(hwnd)


def maximize_window(window_title):
    '''
    Sets focus to window and maximizes window size.

    Parameters
    ----------
    window_title : String
	    Window title.

    Returns
    ----------
    Boolean
	    True for success, False otherwise.
    '''
    hwnd = _get_window_hwnd(window_title)
    if hwnd is not None:
        win32gui.SetForegroundWindow(hwnd)
        win32gui.SetFocus(hwnd)
        return True
    else:
        return False


def verify_focused_window_title(window_title):
    '''
    Verifies that currently focused window contains window_title string.

    Parameters
    ----------
    window_title : String
	    Window title.

    Returns
    ----------
    Boolean
        True if window title contains string window_titl, False if not.
    '''
    hwnd = _get_window_hwnd(window_title)
    hwnd = win32gui.GetForegroundWindow()
    active_window_title = win32gui.GetWindowText(hwnd)
    if window_title.lower().find(active_window_title.lower()) != -1:
        return True
    else:
        return False


def get_window_bbox(window_title):
    '''
    Gets window bounding box.

    Parameters
    ----------
    window_title : String
	    Window title.

    Returns
    ----------
    List
	    Window bounding box as list [x1, y1, x2, y2].
	    None if window title was not found.
    '''
    hwnd = _get_window_hwnd(window_title)
    if hwnd is not None:
        rect = win32gui.GetWindowRect(hwnd)
        return [rect[0], rect[1], rect[2], rect[3]]
    else:
        return None


def _get_window_hwnd(window_title):
    '''
    Gets window hwnd object by Window title.

    Parameters
    ----------
    window_title : String
	    Window title.

    Returns
    ----------
    hwnd Object
        Window hwnd object. None if window title was not found.
    '''
    window_hwnds = _get_windows()
    for hwnd in window_hwnds:
        title = win32gui.GetWindowText(hwnd)
        if window_title.lower().find(title.lower()) != -1:
            return hwnd
    return False


def _get_windows():
    '''
    Gets a list of all windows as hwnd objects.

    Returns
    ----------
    List
	    List of hwnd objects.
    '''
    def callback(hwnd, hwnd_list):
        '''
        Callback function that adds window hwnd to hwnd_list.
        '''
        if win32gui.IsWindowVisible(hwnd):
            window_title = win32gui.GetWindowText(hwnd)
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            if window_title and right-left and bottom-top:
                hwnd_list.append(hwnd)
        return True
    windows_list = []
    win32gui.EnumWindows(callback, windows_list)
    return windows_list
