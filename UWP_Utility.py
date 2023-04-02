#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# https://stackoverflow.com/questions/56861198/how-to-create-python-ctypes-structures-for-ms-windows-package-id-and-package-inf/56892039#56892039
import ctypes
import ctypes.wintypes
import os
import xml.etree.ElementTree as ET
from PyQt5.QtGui import QImage
import win32gui
import win32process

ERROR_SUCCESS = 0x0
ERROR_INSUFFICIENT_BUFFER = 0x7A
APPMODEL_ERROR_NO_PACKAGE = 15700

PACKAGE_FILTER_ALL_LOADED = 0x00000000
PACKAGE_FILTER_HEAD = 0x00000010
PACKAGE_INFORMATION_FULL = 0x00000100
PROCESS_QUERY_LIMITED_INFORMATION = 0x1000

PROCESS_QUERY_INFORMATION = 0x0400
PROCESS_VM_READ =  0x0010


class PACKAGE_INFO_REFERENCE(ctypes.Structure):
    _fields_ = [
        ("reserved", ctypes.c_void_p)
    ]


class PACKAGE_SUBVERSION(ctypes.Structure):

    _fields_ = [
        ("Revision", ctypes.wintypes.USHORT),
        ("Build", ctypes.wintypes.USHORT),
        ("Minor", ctypes.wintypes.USHORT),
        ("Major", ctypes.wintypes.USHORT),
    ]


class PACKAGE_VERSION_U(ctypes.Union):
    _fields_ = [
        ("Version", ctypes.c_uint64),
        ("DUMMYSTRUCTNAME", PACKAGE_SUBVERSION),
    ]


class PACKAGE_VERSION(ctypes.Structure):
    _anonymous_ = ("u", )
    _fields_ = [
        ("u", PACKAGE_VERSION_U),
    ]


class PACKAGE_ID(ctypes.Structure):

    _fields_ = [
        ("reserved", ctypes.c_uint32),
        ("processorArchitecture", ctypes.c_uint32),
        ("version", PACKAGE_VERSION),
        ("name", ctypes.c_wchar_p),
        ("publisher", ctypes.c_wchar_p),
        ("resourceId", ctypes.c_wchar_p),
        ("publisherId", ctypes.c_wchar_p),
    ]


class PACKAGE_INFO(ctypes.Structure):

    _fields_ = [
        ("reserved", ctypes.c_uint32),
        ("flags", ctypes.c_uint32),
        ("path", ctypes.c_wchar_p),
        ("packageFullName", ctypes.c_wchar_p),
        ("packageFamilyName", ctypes.c_wchar_p),
        ("packageId", PACKAGE_ID),
    ]


_user32 = ctypes.WinDLL("user32", use_last_error=True)
_get_windows_thread_process_id = _user32.GetWindowThreadProcessId
_get_windows_thread_process_id.argtypes = (ctypes.wintypes.HWND, ctypes.POINTER(ctypes.wintypes.DWORD))
_get_windows_thread_process_id.restype = ctypes.wintypes.DWORD

WNDENUMPROC  = ctypes.WINFUNCTYPE(ctypes.wintypes.BOOL, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM)

_enum_child_windows = _user32.EnumChildWindows
_enum_child_windows.argtypes = (ctypes.wintypes.HWND, WNDENUMPROC, ctypes.wintypes.LPARAM)
_enum_child_windows.restype = ctypes.wintypes.BOOL

_enum_windows = _user32.EnumWindows
_enum_windows.argtypes = (WNDENUMPROC, ctypes.wintypes.LPARAM)
_enum_windows.restype = ctypes.wintypes.BOOL

_kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
_open_process = _kernel32.OpenProcess
_open_process.argtypes = (ctypes.wintypes.DWORD, ctypes.wintypes.BOOL, ctypes.wintypes.DWORD)
_open_process.restype = ctypes.wintypes.HANDLE

_close_handle = _kernel32.CloseHandle
_close_handle.argtypes = (ctypes.wintypes.HANDLE, )
_close_handle.restype = ctypes.wintypes.BOOL

_get_package_info = _kernel32.GetPackageInfo
_get_package_info.argtypes = (
    PACKAGE_INFO_REFERENCE,
    ctypes.c_uint32,
    ctypes.POINTER(ctypes.c_uint32),
    ctypes.POINTER(ctypes.c_uint8),
    ctypes.POINTER(ctypes.c_uint32)
)
_get_package_info.restype = ctypes.wintypes.LONG

_get_package_full_name = _kernel32.GetPackageFullName
_get_package_full_name.argtypes = (ctypes.wintypes.HANDLE, ctypes.POINTER(ctypes.c_uint32), ctypes.wintypes.LPCWSTR)
_get_package_full_name.restype = ctypes.wintypes.LONG

_get_package_path_by_full_name = _kernel32.GetPackagePathByFullName
_get_package_path_by_full_name.argtypes = (ctypes.wintypes.LPCWSTR, ctypes.POINTER(ctypes.c_uint32), ctypes.wintypes.LPCWSTR)
_get_package_path_by_full_name.restype = ctypes.wintypes.LONG

_package_family_name_from_full_name = _kernel32.PackageFamilyNameFromFullName
_package_family_name_from_full_name.argtypes = (
    ctypes.wintypes.LPCWSTR,
    ctypes.POINTER(ctypes.c_uint32),
    ctypes.wintypes.LPCWSTR)
_package_family_name_from_full_name.restype = ctypes.wintypes.LONG

_open_package_info_by_full_name = _kernel32.OpenPackageInfoByFullName
_open_package_info_by_full_name.argtypes = (
    ctypes.wintypes.LPCWSTR,
    ctypes.c_uint32,
    ctypes.POINTER(PACKAGE_INFO_REFERENCE)
)
_open_package_info_by_full_name.restype = ctypes.wintypes.LONG

_close_package_info = _kernel32.ClosePackageInfo
_close_package_info.argtypes = (
    PACKAGE_INFO_REFERENCE,
)
_close_package_info.restype = ctypes.wintypes.LONG



def get_children(hwnd):
    children = []
    def append_to_collection(element, param):
        children.append(element)
        return True

    func = WNDENUMPROC(append_to_collection)
    _enum_child_windows(hwnd, func, 0)

    return children


def package_full_name_from_handle(handle):
    length = ctypes.c_uint()
    ret_val = _get_package_full_name(handle, ctypes.byref(length), None)
    if ret_val == APPMODEL_ERROR_NO_PACKAGE:
        print(f"package_full_name_from_handle: handle {handle:#x} has no package.")
        return None

    full_name = ctypes.create_unicode_buffer(length.value + 1)
    ret_val = _get_package_full_name(handle, ctypes.byref(length), full_name)
    if ret_val != ERROR_SUCCESS:
        err =  ctypes.WinError(ctypes.get_last_error())
        print(f"package_full_name_from_handle: error -> {str(err)}")
        return None

    return full_name


def package_path_from_full_name(full_name):
    if not full_name:
        raise ValueError("full_name is empty")
        return None
    length = ctypes.c_uint()
    retval = _get_package_path_by_full_name(full_name, ctypes.byref(length), None)
    if retval != ERROR_INSUFFICIENT_BUFFER:
        raise ctypes.WinError(ctypes.get_last_error())

    package_path = ctypes.create_unicode_buffer(length.value)
    retval = _get_package_path_by_full_name(full_name, ctypes.byref(length), package_path)
    if retval != ERROR_SUCCESS:
        raise ctypes.WinError(ctypes.get_last_error())

    if package_path.value == "N/A":
        raise ValueError("Package path is N/A")

    return package_path


def package_family_name_from_full_name(full_name):
    length = ctypes.c_uint()
    retval = _package_family_name_from_full_name(full_name, ctypes.byref(length), None)
    if retval != ERROR_INSUFFICIENT_BUFFER:
        raise ctypes.WinError(ctypes.get_last_error())

    family_name = ctypes.create_unicode_buffer(length.value)
    retval = _package_family_name_from_full_name(full_name, ctypes.byref(length), family_name)
    if retval != ERROR_SUCCESS:
        raise ctypes.WinError(ctypes.get_last_error())

    return family_name


def package_info_reference_from_full_name(full_name):
    package_info_reference = ctypes.pointer(PACKAGE_INFO_REFERENCE())
    retval = _open_package_info_by_full_name(full_name, 0, package_info_reference)
    if retval != ERROR_SUCCESS:
        raise ctypes.WinError(ctypes.get_last_error())

    return package_info_reference


def package_info_buffer_from_reference(package_info_reference):
    length = ctypes.c_uint(0)
    count = ctypes.c_uint()

    retval = _get_package_info(
        package_info_reference.contents,  # package_info_reference is already a pointer. We want its content.
        PACKAGE_FILTER_HEAD,
        ctypes.byref(length),
        None,
        ctypes.byref(count),
    )
    if retval != ERROR_INSUFFICIENT_BUFFER:
        raise ctypes.WinError(ctypes.get_last_error())

    buffer = ctypes.create_string_buffer(length.value)
    buffer_bytes = ctypes.cast(buffer, ctypes.POINTER(ctypes.c_uint8))
    retval = _get_package_info(
        package_info_reference.contents,
        PACKAGE_FILTER_HEAD,
        ctypes.byref(length),
        buffer_bytes,
        ctypes.byref(count),
    )
    if retval != ERROR_SUCCESS:
        raise ctypes.WinError(ctypes.get_last_error())

    return buffer, length


def get_UWP_package_info(hwnd):
    pid = ctypes.wintypes.DWORD()
    _get_windows_thread_process_id(
        hwnd,
        ctypes.byref(pid)
    )

    hprocess = _open_process(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    full_name = package_full_name_from_handle(hprocess)
    if not full_name:
        return
    else:
        print(full_name.value)

    '''
    children = get_children(hwnd)
    for child in children:
        child_pid = ctypes.wintypes.DWORD(0)
        _get_windows_thread_process_id(child, ctypes.byref(child_pid))
        if child_pid != pid:
            hprocess = _open_process(PROCESS_QUERY_LIMITED_INFORMATION, False, child_pid)
            break

    if hprocess is None:
        return

    full_name = package_full_name_from_handle(hprocess)
    if full_name is None:
        return

    if not ("Microsoft.MicrosoftEdge" in full_name.value or "Microsoft.WindowsStore" in full_name.value):
    return None
    '''

    print("=" * 79)
    print("full name: ", full_name.value)
    package_path = package_path_from_full_name(full_name)
    print("package path: ", package_path.value)
    family_name = package_family_name_from_full_name(full_name)
    print("family name:", family_name.value)
    package_info_reference = package_info_reference_from_full_name(full_name)
    print("info reference:", package_info_reference.contents.reserved)

    package_info_buffer, length = package_info_buffer_from_reference(package_info_reference)
    # size_package_info = ctypes.sizeof(PACKAGE_INFO)
    # print(f"PACKAGE_INFO size: {size_package_info:#x}")
    # print(f"num package info: {length.value / size_package_info}")
    package_info = PACKAGE_INFO.from_buffer(package_info_buffer)
    print("packageFullName:", package_info.packageFullName)
    print("=" * 79)

    _close_handle(hprocess)
    _close_package_info(package_info_reference.contents)

def get_windows():
    hwnds = []
    def append_to_collection(element, param):
        hwnds.append(element)
        return True
    func = WNDENUMPROC(append_to_collection)
    _enum_windows(func, 0)
    return hwnds

# 通过 UWP 的包名称获取图标，返回 QImage 对象
def get_icon_from_UWP_package(package_name: str) -> QImage:
    package_path = package_path_from_full_name(package_name)
    if package_path is None:
        raise FileNotFoundError("Package not found")
    package_path = str(package_path.value)
    if package_path is None:
        raise FileNotFoundError("Package not found")
        return None
    # Load the AppxManifest.xml file
    manifest_file = os.path.join(package_path, "AppxManifest.xml")
    if manifest_file is None or not os.path.exists(manifest_file) or not os.access(manifest_file, os.R_OK):
        raise FileNotFoundError("AppxManifest.xml not found")
        return None
    
    tree = ET.parse(manifest_file)
    root = tree.getroot()

    # Find the Logo element
    namespace = root.tag.split('}')[0] + '}'
    logo_element = root.find(namespace + 'Properties').find(namespace + 'Logo')

    # Get the relative path to the logo file
    relative_path = logo_element.text.replace('\\', os.sep)

    # Try to find the exact logo file
    logo_file = os.path.join(package_path, relative_path)
    if os.path.exists(logo_file) and os.access(logo_file, os.R_OK):
        image = QImage(logo_file)
        return image

    # If the exact logo file doesn't exist, try to find a matching file
    else:
        logo_basename = os.path.basename(relative_path)
        logo_dirname = os.path.dirname(relative_path)
        for filename in os.listdir(os.path.join(package_path, logo_dirname)):
            if filename.startswith(logo_basename.split('.')[0]) and filename.endswith('.png'):
                logo_file = os.path.join(package_path, logo_dirname, filename)
                if os.path.exists(logo_file) and os.access(logo_file, os.R_OK):
                    image = QImage(logo_file)
                    return image
                    break
    return None

# 从窗口句柄获取UWP应用的图标
# hwnd: 窗口句柄，但必须是真正的UWP的窗口句柄, 大多数情况下 python 能获取到的只是沙盒的窗口句柄（尽管它们各不相同，但是终会获得一致的沙盒 pid），所以这个函数不可用
def get_icon_from_UWP_hwnd(hwnd: int, image_resize: int = 32) -> QImage:
    pid = ctypes.wintypes.DWORD()
    _get_windows_thread_process_id(
        hwnd,
        ctypes.byref(pid)
    )

    hprocess = _open_process(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    full_name = package_full_name_from_handle(hprocess)
    if not full_name:
        return None
    image = get_icon_from_UWP_package(full_name)
    if image is not None:
        image = image.scaled(image_resize, image_resize)
    return image