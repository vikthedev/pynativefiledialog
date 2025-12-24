#!/usr/bin/env python3
"""
NativeFileDialog — Python wrapper for Windows Vista+ native file dialogs using COM.

Provides a thin, dependency-free interface to Windows File Open / Save dialogs
with support for filters, multi-selection, and folder picking.

Requirements:
    - Windows Vista or newer (Windows NT 6.0+)
"""

__all__ = ("FOS", "NativeFileDialog", "FileFilter", "CommonFilters")

import ctypes
import os
import sys
from ctypes import (
    c_void_p,
    c_ubyte,
    c_ushort,
    c_long,
    byref,
    sizeof,
    cast,
    POINTER,
    WINFUNCTYPE,
)
from ctypes.wintypes import DWORD, HWND, UINT, LPWSTR
from collections.abc import Sequence
from dataclasses import dataclass
from enum import Enum, IntFlag
from uuid import UUID
from typing import Iterable, List, Optional, Tuple, overload, Literal

# ---------------------------------------------------------------------------
# Platform checks
# ---------------------------------------------------------------------------

if os.name != "nt":
    raise OSError("NativeFileDialog is supported on Windows only")

if sys.getwindowsversion().major < 6:
    raise OSError("NativeFileDialog requires Windows Vista or newer (Windows NT 6.0+)")

# ---------------------------------------------------------------------------
# Low-level constants and primitives
# ---------------------------------------------------------------------------

HRESULT = c_long
ERROR_CANCELLED = 0x800704C7
CLSCTX_INPROC_SERVER = 0x1
SIGDN_FILESYSPATH = DWORD(0x80058000)

# ---------------------------------------------------------------------------
# Internal COM pointer helpers
# ---------------------------------------------------------------------------

c_mem_p = POINTER(c_void_p)
PSIZE = sizeof(c_mem_p)

# ---------------------------------------------------------------------------
# Windows API bindings
# ---------------------------------------------------------------------------

ole32 = ctypes.windll.ole32
shell32 = ctypes.windll.shell32

CoCreateInstance = ole32.CoCreateInstance
CoTaskMemFree = ole32.CoTaskMemFree
CoInitialize = ole32.CoInitialize
CoUninitialize = ole32.CoUninitialize
SHCreateItemFromParsingName = shell32.SHCreateItemFromParsingName


# ---------------------------------------------------------------------------
# File filters
# ---------------------------------------------------------------------------

@dataclass(slots=True)
class FileFilter:
    """
    Represents a single file dialog filter.

    Handles extension normalization and produces dialog-ready patterns
    such as "*.png;*.jpg".
    """

    label: str
    _extensions: Tuple[str, ...]

    def __init__(self, label: str, extensions: Iterable[str] = ()):
        self.label = label
        self._extensions = ()
        self.extensions = extensions

    @property
    def extensions(self) -> Tuple[str, ...]:
        """Return normalized extensions."""
        return self._extensions

    @extensions.setter
    def extensions(self, values: Iterable[str]):
        self._extensions = tuple(self._normalize(v) for v in values)
        self.label = self._normalize_label(self.label, self._extensions)

    @property
    def pattern(self) -> str:
        """Return Windows-compatible filter pattern."""
        return "*.*" if not self._extensions else ";".join(self._extensions)

    def matches(self, path: str) -> bool:
        """Check whether the given path matches the filter."""
        if not self._extensions or "*.*" in self._extensions:
            return True
        return path.lower().endswith(tuple(ext.lstrip("*") for ext in self._extensions))

    def normalize_extension(self, path: str) -> str:
        """Append default extension if missing."""
        return path if self.matches(path) else path + self._extensions[0].lstrip("*")

    @classmethod
    def validate(cls, filters: Sequence[Tuple[str, str]]) -> List["FileFilter"]:
        """Convert raw dialog filters into FileFilter objects."""
        return [cls(label, pattern.split(";")) for label, pattern in filters]

    @classmethod
    def prepare(cls, filetypes: Iterable["FileFilter"]) -> Tuple[Tuple[str, str], ...]:
        """Prepare filters for Windows dialog APIs."""
        return tuple((f.label, f.pattern) for f in filetypes if isinstance(f, FileFilter))

    @staticmethod
    def _normalize(ext: str) -> str:
        """Normalize a single file extension."""
        if not ext or ext in ("*", ".", "*.*"):
            ext = "*.*"
        else:
            if ext.endswith("."):
                ext += "*"
            if ext.startswith("."):
                ext = "*" + ext
            if not ext.startswith("*"):
                ext = "*." + ext
        return ext

    @staticmethod
    def _normalize_label(label: str, extensions: tuple) -> str:
        """ Build a user-facing label based on extensions: "<label> (png, jpg)"."""
        if extensions and "*.*" not in extensions:
            cleaned = (ext.lstrip("*.") for ext in extensions)
            label = f"{label} ({', '.join(cleaned)})"
        return label


class CommonFilters(Enum):
    """
    Lazy-initialized collection of commonly used file dialog filters.

    FileFilter objects are only created when accessed via `.filter`.
    """

    # ---------------------------
    # Generic
    # ---------------------------
    ALL = "All files", ("*.*",)
    NO_EXTENSION = "Files without extension", ("*",)

    # ---------------------------
    # Audio
    # ---------------------------
    AUDIO_ALL = "Audio files", ("mp3", "aac", "m4a", "flac", "wav", "w64", "ogg", "opus", "alac", "aiff", "pcm", "raw")
    MP3 = "MP3 audio", ("mp3",)
    AAC = "AAC audio", ("aac", "m4a")
    FLAC = "FLAC audio", ("flac",)
    WAV = "Wave audio", ("wav", "w64")
    OGG = "Ogg Vorbis audio", ("ogg",)
    OPUS = "Opus audio", ("opus",)
    DOLBY = "Dolby lossy audio", ("ac3", "eac3", "ec3")
    AC3 = "AC3 audio", ("ac3",)
    EAC3 = "EAC3 audio", ("eac3", "ec3")
    PCM = "PCM audio", ("pcm", "raw")

    # ---------------------------
    # Video
    # ---------------------------
    VIDEO_ALL = "Video files", ("mp4", "mkv", "avi", "mov", "wmv", "flv", "webm", "mpg", "mpeg", "m4v")
    MP4 = "MP4 video", ("mp4", "m4v")
    MKV = "Matroska video", ("mkv",)
    AVI = "AVI video", ("avi",)
    MOV = "QuickTime video", ("mov",)
    WEBM = "WebM video", ("webm",)
    MEDIA_CONTAINERS = "Media containers", ("mkv", "mka", "mp4", "m4a", "mpa", "avi", "mov")

    # ---------------------------
    # Images
    # ---------------------------
    IMAGE_ALL = "Image files", ("png", "jpg", "jpeg", "bmp", "gif", "tiff", "webp", "heic")
    PNG = "PNG image", ("png",)
    JPEG = "JPEG image", ("jpg", "jpeg")
    BMP = "Bitmap image", ("bmp",)
    GIF = "GIF image", ("gif",)
    TIFF = "TIFF image", ("tiff",)
    WEBP = "WebP image", ("webp",)

    # ---------------------------
    # Documents
    # ---------------------------
    DOCUMENTS = "Documents", ("pdf", "doc", "docx", "xls", "xlsx", "ppt", "pptx", "odt", "ods", "txt", "rtf")
    PDF = "PDF documents", ("pdf",)
    WORD = "Word documents", ("doc", "docx")
    EXCEL = "Excel spreadsheets", ("xls", "xlsx")
    POWERPOINT = "PowerPoint presentations", ("ppt", "pptx")
    TEXT = "Text files", ("txt", "rtf", "md")

    # ---------------------------
    # Archives
    # ---------------------------
    ARCHIVES = "Archive files", ("zip", "rar", "7z", "tar", "gz", "bz2", "xz")
    ZIP = "ZIP archive", ("zip",)
    RAR = "RAR archive", ("rar",)
    SEVEN_Z = "7-Zip archive", ("7z",)
    TAR = "TAR archive", ("tar", "gz", "bz2", "xz")

    # ---------------------------
    # Code / data
    # ---------------------------
    SOURCE_CODE = "Source code", ("py", "c", "cpp", "h", "hpp", "cs", "java", "js", "ts", "rs", "go")
    PYTHON = "Python source", ("py",)
    JSON = "JSON files", ("json",)
    XML = "XML files", ("xml",)
    YAML = "YAML files", ("yml", "yaml")
    CSV = "CSV files", ("csv",)

    # ---------------------------
    # Lazy filter property
    # ---------------------------
    @property
    def filter(self) -> FileFilter:
        """
        Return the FileFilter object, creating it lazily.

        Example usage:
            CommonFilters.PDF.filter
        """
        if not hasattr(self, "_filter"):
            label, ext = self.value
            object.__setattr__(self, "_filter", FileFilter(label, ext))
        return object.__getattribute__(self, "_filter")


# ---------------------------------------------------------------------------
# GUID helpers
# ---------------------------------------------------------------------------

class GUID(ctypes.Structure):
    """
    ctypes-compatible Windows GUID structure.
    """

    _fields_ = (
        ("Data1", DWORD),
        ("Data2", c_ushort),
        ("Data3", c_ushort),
        ("Data4", c_ubyte * 8),
    )

    def __init__(self, value: str | UUID):
        super().__init__()
        u = value if isinstance(value, UUID) else UUID(value)
        self.Data1 = u.time_low
        self.Data2 = u.time_mid
        self.Data3 = u.time_hi_version
        self.Data4[:] = u.bytes[8:]


class _FileDialogGUIDs:
    """
    Internal registry of COM GUIDs used by file dialogs.
    """

    CLSID_FileOpenDialog = GUID("{DC1C5A9C-E88A-4DDE-A5A1-60F82A20AEF7}")
    CLSID_FileSaveDialog = GUID("{C0B4E2F3-BA21-4773-8DBA-335EC946EB8B}")

    IID_IFileOpenDialog = GUID("{D57C7288-D4AD-4768-BE02-9D969532D960}")
    IID_IFileSaveDialog = GUID("{84BCCD23-5FDE-4CDB-AEA4-AF64B83D78AB}")
    IID_IShellItem = GUID("{43826D1E-E718-42EE-BC55-A1E261C37BFE}")


# ---------------------------------------------------------------------------
# COM VTable helpers
# ---------------------------------------------------------------------------

class VTableFunc:
    """
    Helper for accessing COM interface vtable functions.

    VTable indices correspond to definitions in Windows SDK (shobjidl.h).
    """

    FUNCS = {
        "Show": (3, (HWND,)),
        "SetFileType": (4, (UINT, POINTER(LPWSTR * 2))),
        "SetFileTypeIdx": (5, (UINT,)),
        "GetFileTypeIdx": (6, (POINTER(UINT),)),
        "SetOptions": (9, (DWORD,)),
        "GetOptions": (10, (POINTER(DWORD),)),
        "SetFolder": (12, (c_mem_p,)),
        "SetFileName": (15, (LPWSTR,)),
        "SetTitle": (17, (LPWSTR,)),
        "SetOkBtnTxt": (18, (LPWSTR,)),
        "SetFnLabel": (19, (LPWSTR,)),
        "GetResult": (20, (POINTER(c_mem_p),)),
        "GetResults": (27, (POINTER(c_mem_p),)),
        "GetName": (5, (DWORD, POINTER(LPWSTR))),
        "GetCount": (7, (POINTER(DWORD),)),
        "GetItemAt": (8, (DWORD, POINTER(c_mem_p))),
        "Release": (2, ()),
    }

    @classmethod
    def cast(cls, com_obj, name: str):
        """Return a callable COM method from the vtable."""
        index, args = cls.FUNCS[name]
        fn_type = WINFUNCTYPE(HRESULT, c_mem_p, *args)
        return cast(com_obj.contents.value + index * PSIZE, POINTER(fn_type)).contents

    @classmethod
    def free(cls, *com_obj):
        """Release COM objects safely."""
        for obj in com_obj:
            if not cls.is_null_ptr(obj):
                cls.cast(obj, "Release")(obj)

    @staticmethod
    def is_null_ptr(p) -> bool:
        """Check whether a COM pointer is null."""
        return not cast(p, c_void_p).value


# ---------------------------------------------------------------------------
# NativeFileDialog (public API)
# ---------------------------------------------------------------------------

class NativeFileDialog:
    """
    High-level Python wrapper for Windows native File Open / Save dialogs (Vista+).

    Uses the Windows COM-based IFileOpenDialog / IFileSaveDialog APIs and supports:
    - File and folder selection
    - Multi-selection
    - File type filters
    - Custom titles and button labels
    """

    @overload
    @classmethod
    def _get_paths(
            cls,
            title: Optional[str] = None,
            init_dir: Optional[str] = None,
            init_file: Optional[str] = None,
            multichoice: Literal[False] = False,
            confirm_button_label: Optional[str] = None,
            input_label: Optional[str] = None,
            add_flags: int = 0,
            file_type_filters: Sequence[FileFilter] = (),
            save_mode: bool = False,
    ) -> Optional[str]:
        ...

    @overload
    @classmethod
    def _get_paths(
            cls,
            title: Optional[str] = None,
            init_dir: Optional[str] = None,
            init_file: Optional[str] = None,
            multichoice: Literal[True] = True,
            confirm_button_label: Optional[str] = None,
            input_label: Optional[str] = None,
            add_flags: int = 0,
            file_type_filters: Sequence[FileFilter] = (),
            save_mode: bool = False,
    ) -> Optional[List[str]]:
        ...

    @classmethod
    def _get_paths(
            cls,
            title: Optional[str] = None,
            init_dir: Optional[str] = None,
            init_file: Optional[str] = None,
            multichoice: bool = False,
            confirm_button_label: Optional[str] = None,
            input_label: Optional[str] = None,
            add_flags: int = 0,
            file_type_filters: Sequence[FileFilter] = (),
            save_mode: bool = False,
    ):
        """
        Internal dialog engine.

        Handles COM initialization, dialog configuration, execution,
        and result extraction.

        Returns:
            str | List[str] | None:
                - str        if multichoice is False
                - List[str]  if multichoice is True
                - None       if dialog was cancelled
        """

        # HWND to attach the dialog to (0 = no owner window)
        window_id: int = 0

        flags = DWORD()
        COM, DIR, item, mult = c_mem_p(), c_mem_p(), c_mem_p(), c_mem_p()

        paths: List[str] = []
        initialized = False

        try:
            # Initialize COM for the current thread
            if CoInitialize(None) < 0:
                raise OSError("CoInitialize failed")

            initialized = True

            # Create File Open or File Save dialog COM object
            if (
                    CoCreateInstance(
                        byref(
                            _FileDialogGUIDs.CLSID_FileSaveDialog
                            if save_mode
                            else _FileDialogGUIDs.CLSID_FileOpenDialog
                        ),
                        None,
                        CLSCTX_INPROC_SERVER,
                        byref(
                            _FileDialogGUIDs.IID_IFileSaveDialog
                            if save_mode
                            else _FileDialogGUIDs.IID_IFileOpenDialog
                        ),
                        byref(COM),
                    ) < 0
                    or VTableFunc.is_null_ptr(COM)
            ):
                raise OSError("CoCreateInstance failed")

            # Resolve required COM vtable functions
            Show = VTableFunc.cast(COM, "Show")
            SetFileType = VTableFunc.cast(COM, "SetFileType")
            SetFileTypeIdx = VTableFunc.cast(COM, "SetFileTypeIdx")
            GetFileTypeIdx = VTableFunc.cast(COM, "GetFileTypeIdx")
            SetOptions = VTableFunc.cast(COM, "SetOptions")
            GetOptions = VTableFunc.cast(COM, "GetOptions")
            SetFolder = VTableFunc.cast(COM, "SetFolder")
            SetFileName = VTableFunc.cast(COM, "SetFileName")
            SetTitle = VTableFunc.cast(COM, "SetTitle")
            SetOkBtnTxt = VTableFunc.cast(COM, "SetOkBtnTxt")
            SetFnLabel = VTableFunc.cast(COM, "SetFnLabel")
            GetResult = VTableFunc.cast(COM, "GetResult")
            GetResults = VTableFunc.cast(COM, "GetResults")

            # Configure dialog flags
            GetOptions(COM, byref(flags))
            flags.value |= (
                    FOS.FORCEFILESYSTEM
                    | FOS.PATHMUSTEXIST
                    | FOS.FILEMUSTEXIST
                    | (bool(multichoice) and FOS.ALLOWMULTISELECT)
                    | add_flags
            )
            SetOptions(COM, flags)

            # Configure file type filters
            if prepared_filters := FileFilter.prepare(file_type_filters):
                SetFileType(
                    COM,
                    len(prepared_filters),
                    (LPWSTR * 2 * len(prepared_filters))(
                        *[tuple(LPWSTR(i) for i in j) for j in prepared_filters]
                    ),
                )
                SetFileTypeIdx(COM, 1)

            # Optional dialog customization
            if title:
                SetTitle(COM, LPWSTR(title))
            if init_file:
                SetFileName(COM, LPWSTR(init_file))
            if confirm_button_label:
                SetOkBtnTxt(COM, LPWSTR(confirm_button_label))
            if input_label:
                SetFnLabel(COM, LPWSTR(input_label))

            # Set initial directory if provided
            if (init_dir and
                    SHCreateItemFromParsingName(
                        LPWSTR(init_dir), None, byref(_FileDialogGUIDs.IID_IShellItem), byref(DIR)
                    ) >= 0):
                SetFolder(COM, DIR)

            # Show the dialog
            hr = Show(COM, HWND(window_id))

            if hr != ctypes.c_long(ERROR_CANCELLED).value:
                if hr < 0:
                    raise OSError(f"Dialog failed with HRESULT {hr:#x}")

                # Retrieve result(s)
                path = LPWSTR()

                if save_mode:
                    # Single file result (Save dialog)
                    if GetResult(COM, byref(item)) >= 0:
                        GetName = VTableFunc.cast(item, "GetName")
                        Release = VTableFunc.cast(item, "Release")

                        if GetName(item, SIGDN_FILESYSPATH, byref(path)) >= (
                                Release(item) and 0
                        ):
                            GetFileTypeIdx(COM, byref(filetypeidx := UINT()))
                            try:
                                raw_filter = prepared_filters[filetypeidx.value - 1]
                                ff = FileFilter.validate((raw_filter,))[0]
                                paths.append(ff.normalize_extension(path.value))
                            except (IndexError, TypeError, AttributeError):
                                paths.append(path.value)

                            CoTaskMemFree(path)

                else:
                    # File Open dialog (single or multi-selection)
                    if GetResults(COM, byref(mult)) >= 0:
                        GetCount = VTableFunc.cast(mult, "GetCount")
                        GetItemAt = VTableFunc.cast(mult, "GetItemAt")

                        pathsnum = DWORD()
                        if GetCount(mult, byref(pathsnum)) >= 0:
                            for i in range(pathsnum.value):
                                if GetItemAt(mult, i, byref(item)) < 0:
                                    break

                                # Resolve per-item methods lazily
                                GetName = cast(
                                    item.contents.value + 5 * PSIZE,
                                    POINTER(
                                        WINFUNCTYPE(
                                            HRESULT, c_mem_p, DWORD, POINTER(LPWSTR)
                                        )
                                    ),
                                ).contents
                                Release = cast(
                                    item.contents.value + 2 * PSIZE,
                                    POINTER(WINFUNCTYPE(HRESULT, c_mem_p)),
                                ).contents

                                if GetName(item, SIGDN_FILESYSPATH, byref(path)) < (
                                        Release(item) and 0
                                ):
                                    break

                                paths.append(path.value)
                                CoTaskMemFree(path)

        finally:
            # Release all allocated COM objects and uninitialize COM
            VTableFunc.free(item, mult, DIR, COM)
            if initialized:
                CoUninitialize()

        result = paths
        if result and not multichoice:
            result = result[0]

        return result or None

    # ---------------------------
    # Public API
    # ---------------------------

    @classmethod
    def get_dir(
            cls,
            title: Optional[str] = None,
            init_dir: Optional[str] = None,
            confirm_button_label: Optional[str] = None,
            input_label: Optional[str] = None,
            flags: int = 0,
    ) -> Optional[str]:
        """
        Open a native folder selection dialog.

        Returns:
            Selected directory path or None if cancelled.
        """
        return cls._get_paths(
            title=title,
            init_dir=init_dir,
            confirm_button_label=confirm_button_label,
            input_label=input_label,
            add_flags=FOS.PICKFOLDERS | flags,
        )

    @classmethod
    def get_file(
            cls,
            title: Optional[str] = None,
            init_dir: Optional[str] = None,
            init_file: Optional[str] = None,
            file_type_filters: Sequence[FileFilter] = (
                    CommonFilters.AUDIO_ALL.filter,
                    CommonFilters.MEDIA_CONTAINERS.filter,
                    CommonFilters.ALL.filter
            ),
            confirm_button_label: Optional[str] = None,
            input_label: Optional[str] = None,
            flags: int = 0,
    ) -> Optional[str]:
        """
        Open a native single-file selection dialog.

        Returns:
            Selected file path or None if cancelled.
        """
        return cls._get_paths(
            title=title,
            init_dir=init_dir,
            init_file=init_file,
            confirm_button_label=confirm_button_label,
            input_label=input_label,
            add_flags=flags,
            file_type_filters=file_type_filters,
        )

    @classmethod
    def get_files(
            cls,
            title: Optional[str] = None,
            init_dir: Optional[str] = None,
            init_file: Optional[str] = None,
            file_type_filters: Sequence[FileFilter] = (
                    CommonFilters.AUDIO_ALL.filter,
                    CommonFilters.MEDIA_CONTAINERS.filter,
                    CommonFilters.ALL.filter
            ),
            confirm_button_label: Optional[str] = None,
            input_label: Optional[str] = None,
            flags: int = 0,
    ) -> Optional[List[str]]:
        """
        Open a native multi-file selection dialog.

        Returns:
            List of selected file paths or None if cancelled.
        """
        return cls._get_paths(
            title=title,
            init_dir=init_dir,
            init_file=init_file,
            multichoice=True,
            confirm_button_label=confirm_button_label,
            input_label=input_label,
            add_flags=flags,
            file_type_filters=file_type_filters,
        )

    @classmethod
    def set_file(
            cls,
            title: Optional[str] = None,
            init_dir: Optional[str] = None,
            init_file: Optional[str] = None,
            file_type_filters: Sequence[FileFilter] = (
                    CommonFilters.AUDIO_ALL.filter,
                    CommonFilters.ALL.filter
            ),
            confirm_button_label: Optional[str] = None,
            input_label: Optional[str] = None,
            flags: int = 0,
    ) -> Optional[str]:
        """
        Open a native file save dialog.

        Returns:
            Selected output file path or None if cancelled.
        """
        return cls._get_paths(
            title=title,
            init_dir=init_dir,
            init_file=init_file,
            confirm_button_label=confirm_button_label,
            input_label=input_label,
            add_flags=(bool(file_type_filters) and FOS.STRICTFILETYPES) | flags,
            file_type_filters=file_type_filters,
            save_mode=True,
        )


# ---------------------------------------------------------------------------
# Flags / enums
# ---------------------------------------------------------------------------

class FOS(IntFlag):
    """
    File Open / Save dialog option flags (FOS_*).

    See:
        https://learn.microsoft.com/en-us/windows/win32/api/shobjidl_core/ne-shobjidl_core-fileopendialogoptions
    """

    OVERWRITEPROMPT = 0x2
    STRICTFILETYPES = 0x4
    NOCHANGEDIR = 0x8
    PICKFOLDERS = 0x20
    FORCEFILESYSTEM = 0x40
    ALLNONSTORAGEITEMS = 0x80
    NOVALIDATE = 0x100
    ALLOWMULTISELECT = 0x200
    PATHMUSTEXIST = 0x800
    FILEMUSTEXIST = 0x1000
    CREATEPROMPT = 0x2000
    SHAREAWARE = 0x4000
    NOREADONLYRETURN = 0x8000
    NOTESTFILECREATE = 0x10000
    HIDEMRUPLACES = 0x20000
    HIDEPINNEDPLACES = 0x40000
    NODEREFERENCELINKS = 0x100000
    OKBUTTONNEEDSINTERACTION = 0x200000
    DONTADDTORECENT = 0x2000000
    FORCESHOWHIDDEN = 0x10000000
    DEFAULTNOMINIMODE = 0x20000000
    FORCEPREVIEWPANEON = 0x40000000
    SUPPORTSTREAMABLEITEMS = 0x80000000
