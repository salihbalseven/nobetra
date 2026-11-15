"""Keep owned dialogs visible and reopen the existing Windows session."""

from contextlib import contextmanager
import hashlib
import os
import time
import tkinter as tk


def install(root):
    """Restore an owned Tk dialog when the main window is mapped or activated."""
    pending = None

    def restore():
        nonlocal pending
        pending = None
        if root.state() not in ('normal', 'zoomed'):
            return
        grabbed = root.grab_current()
        if grabbed is None:
            return
        dialog = grabbed.winfo_toplevel()
        if dialog is root or not getattr(dialog, '_nobet_modal', False):
            return
        if dialog.state() in ('iconic', 'withdrawn'):
            dialog.deiconify()
        dialog.lift()
        focus = root.focus_get()
        if focus is None or focus.winfo_toplevel() is root:
            dialog.focus_set()

    def activated(event):
        nonlocal pending
        if event.widget is root and pending is None:
            pending = root.after_idle(restore)

    root.bind('<Map>', activated, add='+')
    root.bind('<FocusIn>', activated, add='+')


def modal(window, owner):
    """Give a dialog an owner and return focus to the previous dialog on close."""
    owner = owner.winfo_toplevel()
    prior = owner.grab_current()
    window.transient(owner)
    window._nobet_modal = True
    window.grab_set()

    def show(event):
        if event.widget is window and owner.state() in ('normal', 'zoomed'):
            window.lift()

    def close(event):
        if event.widget is not window or prior is None:
            return
        def restore_previous():
            if prior.winfo_exists() and prior.winfo_viewable() and prior.grab_current() is None:
                prior.grab_set()
                prior.winfo_toplevel().lift()
        owner.after_idle(restore_previous)

    window.bind('<Map>', show, add='+')
    window.bind('<Destroy>', close, add='+')


def window_key(database):
    return 'Nobetra_' + hashlib.sha256(str(database.resolve()).casefold().encode()).hexdigest()


def register(root, database):
    if os.name != 'nt':
        return
    import ctypes
    from ctypes import wintypes

    user = ctypes.WinDLL('user32', use_last_error=True)
    user.GetAncestor.argtypes = (wintypes.HWND, wintypes.UINT)
    user.GetAncestor.restype = wintypes.HWND
    user.SetPropW.argtypes = (wintypes.HWND, wintypes.LPCWSTR, wintypes.HANDLE)
    root.update_idletasks()
    client = root.winfo_id()
    window = user.GetAncestor(client, 2) or client  # GA_ROOT
    if not user.SetPropW(window, window_key(database), 1):
        raise ctypes.WinError(ctypes.get_last_error())


@contextmanager
def single_instance(database):
    """Open one app window per database and activate it on a second launch."""
    if os.name != 'nt':
        yield True
        return

    import ctypes
    from ctypes import wintypes

    kernel = ctypes.WinDLL('kernel32', use_last_error=True)
    user = ctypes.WinDLL('user32', use_last_error=True)
    kernel.CreateMutexW.argtypes = (ctypes.c_void_p, wintypes.BOOL, wintypes.LPCWSTR)
    kernel.CreateMutexW.restype = wintypes.HANDLE
    kernel.CloseHandle.argtypes = (wintypes.HANDLE,)
    callback = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    user.EnumWindows.argtypes = (callback, wintypes.LPARAM)
    user.GetPropW.argtypes = (wintypes.HWND, wintypes.LPCWSTR)
    user.GetPropW.restype = wintypes.HANDLE
    user.ShowWindowAsync.argtypes = (wintypes.HWND, ctypes.c_int)
    user.IsIconic.argtypes = (wintypes.HWND,)
    user.GetLastActivePopup.argtypes = (wintypes.HWND,)
    user.GetLastActivePopup.restype = wintypes.HWND
    user.SetForegroundWindow.argtypes = (wintypes.HWND,)

    identity = window_key(database)
    mutex = kernel.CreateMutexW(None, False, 'Local\\' + identity)
    error = ctypes.get_last_error()
    if not mutex:
        raise ctypes.WinError(error)
    try:
        if error != 183:
            yield True
            return
        hwnd = None
        @callback
        def find(window, _):
            nonlocal hwnd
            if user.GetPropW(window, identity):
                hwnd = window
                return False
            return True

        for _ in range(30):
            user.EnumWindows(find, 0)
            if hwnd:
                if user.IsIconic(hwnd):
                    user.ShowWindowAsync(hwnd, 9)  # SW_RESTORE
                target = user.GetLastActivePopup(hwnd) or hwnd
                user.SetForegroundWindow(target)
                break
            time.sleep(.1)
        yield False
    finally:
        kernel.CloseHandle(mutex)
