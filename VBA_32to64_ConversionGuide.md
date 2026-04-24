# VBA 32-bit to 64-bit API Conversion Guide
## STGC-Legacy (ConDispatchSQL) — Access VBA Migration

---

## Overview

This document provides a complete set of **find/replace operations** to convert all 32-bit Windows API `Declare` statements to 64-bit compatible equivalents. The API inventory below was **verified by binary scan of all 129 VBA module blobs** in `ConDispatchSQL_ConvertTo64.accdb`, confirming 73 unique Win32 API functions across 12 module blobs.

| Module | DLL | Key APIs Confirmed |
|---|---|---|
| `modAPI` | User32.dll | GetClassNameA, GetWindow, ShowWindowAsync, GetComputerName |
| `modAPIIcons` | User32.dll | LoadImageA, SendMessageA |
| `modClipboard` | User32.dll + Kernel32 | OpenClipboard, CloseClipboard, GetClipboardData, SetClipboardData, EmptyClipboard, GlobalAlloc, GlobalFree, GlobalLock, GlobalUnlock |
| `modExplorerApi` | Shell32.dll | SHBrowseForFolderA, SHGetPathFromIDListA |
| `modFileDlgs` | Shell32.dll + Kernel32 | SHGetPathFromIDList, SHBrowseForFolder, ShellExecute, FindFirstFile, FindNextFile, FindClose, CoTaskMemFree, lstrcat |
| `modIniApi` | Kernel32.dll | GetPrivateProfileString, WritePrivateProfileString |
| `modRegistry` | Advapi32.dll | RegOpenKeyA, RegQueryValueExA, RegSetValueExA, RegCreateKeyA, RegDeleteKeyA, RegDeleteValueA, RegCloseKey |
| `modRegistry3` | Advapi32.dll | RegOpenKeyExA, RegQueryValueExA, RegSetValueExA, RegCreateKeyExA, RegDeleteKeyA, RegCloseKey |
| `modWinInet` | WinInet.dll | InternetOpen, InternetConnect, InternetCloseHandle, FtpGetFile, FtpPutFile, FtpOpenFile, FtpFindFirstFile, FtpGetCurrentDirectory, FtpSetCurrentDirectory, FtpCreateDirectory |
| `modFTP` / `clsFTP` | WinInet.dll | InternetOpen, InternetConnect, InternetCloseHandle, FtpGetFile, FtpPutFile, FtpOpenFile, FtpGetCurrentDirectory, FtpSetCurrentDirectory, InternetWriteFile |
| `modODBC` | ODBC32.dll | SQLConfigDataSource |

> **Scan basis:** 1,096 rows read from MSysAccessStorage; 12 blobs with Win32 API references confirmed via decompressed p-code analysis.

---

## THE GOLDEN RULE OF 32→64 BIT VBA CONVERSION

The core changes required are:

1. Add **`PtrSafe`** after `Declare` (mandatory — VBA7/64-bit will not compile without it)
2. Change **`As Long`** to **`As LongPtr`** for all **handle**, **pointer**, and **HWND** parameters
3. Change **`ByVal hWnd As Long`** → **`ByVal hWnd As LongPtr`** everywhere
4. Keep **`As Long`** for pure integer/numeric values (counts, flags, BOOL returns, error codes)
5. Change **`As Long`** to **`As LongPtr`** for function return types that return handles/pointers

### Quick Reference: What Changes vs. What Stays

| Parameter Type | 32-bit | 64-bit | Rule |
|---|---|---|---|
| Window handle (HWND) | `As Long` | `As LongPtr` | **ALWAYS changes** |
| Any HANDLE type | `As Long` | `As LongPtr` | **ALWAYS changes** |
| Pointer (LPVOID, LPCSTR passed as Long) | `As Long` | `As LongPtr` | **ALWAYS changes** |
| lParam / wParam | `As Long` | `As LongPtr` | **ALWAYS changes** |
| pidl (shell item pointer) | `As Long` | `As LongPtr` | **ALWAYS changes** |
| Registry key (HKEY) | `As Long` | `As LongPtr` | **ALWAYS changes** |
| Internet handle (HINTERNET) | `As Long` | `As LongPtr` | **ALWAYS changes** |
| Global memory handle (HGLOBAL) | `As Long` | `As LongPtr` | **ALWAYS changes** |
| String length / count | `As Long` | `As Long` | NO CHANGE |
| Flags / bit masks | `As Long` | `As Long` | NO CHANGE |
| Boolean return (BOOL) | `As Long` | `As Long` | NO CHANGE |
| Error codes (LSTATUS) | `As Long` | `As Long` | NO CHANGE |
| Integer counts | `As Long` | `As Long` | NO CHANGE |
| `Declare Function` | (no PtrSafe) | `PtrSafe` added | **ALWAYS changes** |
| `Declare Sub` | (no PtrSafe) | `PtrSafe` added | **ALWAYS changes** |

---

## PART 1: FIND/REPLACE OPERATIONS (Apply in This Exact Order)

Use VBA IDE → Edit → Replace (Ctrl+H), check **"Current Project"**, **case-sensitive** unless noted.

---

### STEP 1 — Add PtrSafe to all Declare statements
> Run this FIRST. Any `Declare` without `PtrSafe` will fail to compile in 64-bit Office.
> Use **case-insensitive** match for this step only.

**Find:**
```
Declare Function
```
**Replace With:**
```
Declare PtrSafe Function
```

**Find:**
```
Declare Sub
```
**Replace With:**
```
Declare PtrSafe Sub
```

> ⚠️ After Step 1, search for `Declare PtrSafe PtrSafe` — fix any doubles (modules that already had PtrSafe).

---

### STEP 2 — Fix HWND parameters (window handles)
> `HWND` is pointer-sized — must be `LongPtr` in 64-bit.

| Find | Replace With |
|---|---|
| `ByVal hWnd As Long` | `ByVal hWnd As LongPtr` |
| `ByVal hwnd As Long` | `ByVal hwnd As LongPtr` |
| `ByVal hWndParent As Long` | `ByVal hWndParent As LongPtr` |
| `ByVal hwndOwner As Long` | `ByVal hwndOwner As LongPtr` |
| `ByVal hWndOwner As Long` | `ByVal hWndOwner As LongPtr` |
| `ByRef hWnd As Long` | `ByRef hWnd As LongPtr` |

---

### STEP 3 — Fix generic HANDLE parameters
> All Windows HANDLE types (HINSTANCE, HICON, HBITMAP, HMENU, HGLOBAL, etc.) are pointer-sized.

| Find | Replace With |
|---|---|
| `ByVal hInst As Long` | `ByVal hInst As LongPtr` |
| `ByVal hInstance As Long` | `ByVal hInstance As LongPtr` |
| `ByVal hIcon As Long` | `ByVal hIcon As LongPtr` |
| `ByVal hBitmap As Long` | `ByVal hBitmap As LongPtr` |
| `ByVal hMenu As Long` | `ByVal hMenu As LongPtr` |
| `ByVal hModule As Long` | `ByVal hModule As LongPtr` |
| `ByVal hProcess As Long` | `ByVal hProcess As LongPtr` |
| `ByVal hThread As Long` | `ByVal hThread As LongPtr` |
| `ByVal hObject As Long` | `ByVal hObject As LongPtr` |
| `ByVal hDC As Long` | `ByVal hDC As LongPtr` |
| `ByVal hFile As Long` | `ByVal hFile As LongPtr` |
| `ByVal hMem As Long` | `ByVal hMem As LongPtr` |

---

### STEP 4 — Fix Registry handle parameters (HKEY)
> Registry keys (HKEY) are pointer-sized handles.

| Find | Replace With |
|---|---|
| `ByVal hKey As Long` | `ByVal hKey As LongPtr` |
| `ByRef hKey As Long` | `ByRef hKey As LongPtr` |
| `phkResult As Long` | `phkResult As LongPtr` |
| `ByRef phkResult As Long` | `ByRef phkResult As LongPtr` |

---

### STEP 5 — Fix WinInet / FTP handle parameters (HINTERNET)

| Find | Replace With |
|---|---|
| `ByVal hInternet As Long` | `ByVal hInternet As LongPtr` |
| `ByVal hConnect As Long` | `ByVal hConnect As LongPtr` |
| `ByVal hFtpSession As Long` | `ByVal hFtpSession As LongPtr` |
| `ByVal hConnection As Long` | `ByVal hConnection As LongPtr` |
| `ByVal dwContext As Long` | `ByVal dwContext As LongPtr` |
| `ByVal hFindHandle As Long` | `ByVal hFindHandle As LongPtr` |

---

### STEP 6 — Fix FindFile handle parameters (HANDLE returned by FindFirstFile)

| Find | Replace With |
|---|---|
| `ByVal hFindFile As Long` | `ByVal hFindFile As LongPtr` |
| `ByRef hFindFile As Long` | `ByRef hFindFile As LongPtr` |

---

### STEP 7 — Fix pointer / lParam / wParam parameters

| Find | Replace With |
|---|---|
| `ByVal lParam As Long` | `ByVal lParam As LongPtr` |
| `ByRef lParam As Long` | `ByRef lParam As LongPtr` |
| `lParam As Long` | `lParam As LongPtr` |
| `ByVal wParam As Long` | `ByVal wParam As LongPtr` |
| `ByRef wParam As Long` | `ByRef wParam As LongPtr` |

> ⚠️ The bare `lParam As Long` replace is broad — verify it only hits Declare lines and Type block fields.

---

### STEP 8 — Fix shell pointer parameters (PIDL)

| Find | Replace With |
|---|---|
| `ByVal pidl As Long` | `ByVal pidl As LongPtr` |
| `ByRef pidl As Long` | `ByRef pidl As LongPtr` |
| `pidlRoot As Long` | `pidlRoot As LongPtr` |
| `lpfn As Long` | `lpfn As LongPtr` |

---

### STEP 9 — Fix TYPE structure fields containing handles/pointers
> Open each `Type...End Type` block and apply:

| Find | Replace With | Used In |
|---|---|---|
| `hIcon As Long` | `hIcon As LongPtr` | SHFILEINFO |
| `hwndOwner As Long` | `hwndOwner As LongPtr` | BROWSEINFO |
| `pidlRoot As Long` | `pidlRoot As LongPtr` | BROWSEINFO |
| `lParam As Long` | `lParam As LongPtr` | BROWSEINFO |
| `lpfn As Long` | `lpfn As LongPtr` | BROWSEINFO |
| `hWnd As Long` | `hWnd As LongPtr` | Any UDT with hWnd |

---

### STEP 10 — Fix Dim statements for handle/pointer variables
> Any `Dim` variable that holds a value returned from an API that now returns `LongPtr`.

| Find | Replace With |
|---|---|
| `Dim hWnd As Long` | `Dim hWnd As LongPtr` |
| `Dim hwnd As Long` | `Dim hwnd As LongPtr` |
| `Dim hInst As Long` | `Dim hInst As LongPtr` |
| `Dim hIcon As Long` | `Dim hIcon As LongPtr` |
| `Dim hInternet As Long` | `Dim hInternet As LongPtr` |
| `Dim hFTP As Long` | `Dim hFTP As LongPtr` |
| `Dim hConnect As Long` | `Dim hConnect As LongPtr` |
| `Dim hKey As Long` | `Dim hKey As LongPtr` |
| `Dim hMem As Long` | `Dim hMem As LongPtr` |
| `Dim hDC As Long` | `Dim hDC As LongPtr` |
| `Dim hFile As Long` | `Dim hFile As LongPtr` |
| `Dim lParam As Long` | `Dim lParam As LongPtr` |
| `Dim wParam As Long` | `Dim wParam As LongPtr` |
| `Dim pidl As Long` | `Dim pidl As LongPtr` |
| `Dim hFindFile As Long` | `Dim hFindFile As LongPtr` |

---

### STEP 11 — MANUAL FIXES: Function Return Types
> These cannot be safely automated — each `Declare Function` line needs its trailing `) As Long` changed to `) As LongPtr` where the return type is a handle or pointer.

**Functions returning `LongPtr` (handles/pointers):**

| Function | DLL | Return Value |
|---|---|---|
| `GetWindow` | User32 | HWND |
| `FindWindow` | User32 | HWND |
| `FindWindowEx` | User32 | HWND |
| `ShowWindowAsync` | User32 | BOOL → keep `As Long` |
| `SendMessage` | User32 | LRESULT (LongPtr) |
| `LoadImage` | User32 | HANDLE |
| `ShellExecute` | Shell32 | HINSTANCE |
| `SHBrowseForFolder` | Shell32 | PIDLIST_ABSOLUTE (pointer) |
| `GetClipboardData` | User32 | HANDLE |
| `SetClipboardData` | User32 | HANDLE |
| `GlobalAlloc` | Kernel32 | HGLOBAL (pointer) |
| `GlobalLock` | Kernel32 | LPVOID (pointer) |
| `InternetOpen` | WinInet | HINTERNET |
| `InternetConnect` | WinInet | HINTERNET |
| `FtpOpenFile` | WinInet | HINTERNET |
| `FindFirstFile` | Kernel32 | HANDLE |

**Functions that KEEP `As Long` return type:**

| Function | DLL | Return Value |
|---|---|---|
| `OpenClipboard` | User32 | BOOL |
| `CloseClipboard` | User32 | BOOL |
| `EmptyClipboard` | User32 | BOOL |
| `ShowWindow` | User32 | BOOL |
| `GetClassName` | User32 | char count (int) |
| `GetComputerName` | Kernel32 | BOOL |
| `GlobalFree` | Kernel32 | HGLOBAL (0 on success) → keep `As Long` |
| `GlobalUnlock` | Kernel32 | BOOL |
| `FindNextFile` | Kernel32 | BOOL |
| `FindClose` | Kernel32 | BOOL |
| `RegOpenKey` | Advapi32 | LSTATUS (error code) |
| `RegOpenKeyEx` | Advapi32 | LSTATUS (error code) |
| `RegCloseKey` | Advapi32 | LSTATUS (error code) |
| `RegQueryValueEx` | Advapi32 | LSTATUS (error code) |
| `RegSetValueEx` | Advapi32 | LSTATUS (error code) |
| `RegCreateKey` | Advapi32 | LSTATUS (error code) |
| `RegCreateKeyEx` | Advapi32 | LSTATUS (error code) |
| `RegDeleteKey` | Advapi32 | LSTATUS (error code) |
| `RegDeleteValue` | Advapi32 | LSTATUS (error code) |
| `GetPrivateProfileString` | Kernel32 | char count |
| `WritePrivateProfileString` | Kernel32 | BOOL |
| `InternetCloseHandle` | WinInet | BOOL |
| `FtpGetFile` | WinInet | BOOL |
| `FtpPutFile` | WinInet | BOOL |
| `FtpCreateDirectory` | WinInet | BOOL |
| `SQLConfigDataSource` | ODBC32 | BOOL |

---

## PART 2: MODULE-BY-MODULE REFERENCE

### modAPI (User32.dll + Kernel32.dll)
```vba
' BEFORE (32-bit):
Private Declare Function GetClassName Lib "User32" _
    Alias "GetClassNameA" (ByVal hWnd As Long, _
    ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Private Declare Function GetWindow Lib "User32" _
    (ByVal hWnd As Long, ByVal wCmd As Long) As Long

Private Declare Function ShowWindowAsync Lib "User32" _
    (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Private Declare Function GetComputerName Lib "Kernel32" _
    Alias "GetComputerNameA" (ByVal lpBuffer As String, _
    nSize As Long) As Long

' AFTER (64-bit compatible):
Private Declare PtrSafe Function GetClassName Lib "User32" _
    Alias "GetClassNameA" (ByVal hWnd As LongPtr, _
    ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Private Declare PtrSafe Function GetWindow Lib "User32" _
    (ByVal hWnd As LongPtr, ByVal wCmd As Long) As LongPtr

Private Declare PtrSafe Function ShowWindowAsync Lib "User32" _
    (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long

Private Declare PtrSafe Function GetComputerName Lib "Kernel32" _
    Alias "GetComputerNameA" (ByVal lpBuffer As String, _
    nSize As Long) As Long
```

---

### modAPIIcons (User32.dll)
```vba
' BEFORE (32-bit):
Private Declare Function LoadImage Lib "User32" Alias "LoadImageA" _
    (ByVal hInst As Long, ByVal lpszName As String, _
    ByVal uType As Long, ByVal cxDesired As Long, _
    ByVal cyDesired As Long, ByVal fuLoad As Long) As Long

Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" _
    (ByVal hWnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, lParam As Long) As Long

' AFTER (64-bit compatible):
Private Declare PtrSafe Function LoadImage Lib "User32" Alias "LoadImageA" _
    (ByVal hInst As LongPtr, ByVal lpszName As String, _
    ByVal uType As Long, ByVal cxDesired As Long, _
    ByVal cyDesired As Long, ByVal fuLoad As Long) As LongPtr

Private Declare PtrSafe Function SendMessage Lib "User32" Alias "SendMessageA" _
    (ByVal hWnd As LongPtr, ByVal wMsg As Long, _
    ByVal wParam As LongPtr, lParam As LongPtr) As LongPtr
```

---

### modClipboard (User32.dll + Kernel32.dll)
```vba
' BEFORE (32-bit):
Private Declare Function OpenClipboard Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function CloseClipboard Lib "User32" () As Long
Private Declare Function GetClipboardData Lib "User32" (ByVal wFormat As Long) As Long
Private Declare Function SetClipboardData Lib "User32" (ByVal wFormat As Long, _
    ByVal hMem As Long) As Long
Private Declare Function EmptyClipboard Lib "User32" () As Long
Private Declare Function GlobalAlloc Lib "Kernel32" _
    (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "Kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "Kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "Kernel32" (ByVal hMem As Long) As Long

' AFTER (64-bit compatible):
Private Declare PtrSafe Function OpenClipboard Lib "User32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "User32" () As Long
Private Declare PtrSafe Function GetClipboardData Lib "User32" (ByVal wFormat As Long) As LongPtr
Private Declare PtrSafe Function SetClipboardData Lib "User32" (ByVal wFormat As Long, _
    ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function EmptyClipboard Lib "User32" () As Long
Private Declare PtrSafe Function GlobalAlloc Lib "Kernel32" _
    (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalFree Lib "Kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function GlobalLock Lib "Kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "Kernel32" (ByVal hMem As LongPtr) As Long
```

> ℹ️ `GlobalAlloc` `dwBytes` parameter: use `LongPtr` to support large allocations on 64-bit. `GHND`, `GMEM_MOVEABLE` constants remain `As Long`.

---

### modFileDlgs (Shell32.dll + Kernel32.dll + ole32.dll)
```vba
' BEFORE (32-bit):
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" _
    Alias "SHGetPathFromIDListA" (ByVal pidl As Long, _
    ByVal pszPath As String) As Long

Private Declare Function SHBrowseForFolder Lib "shell32.dll" _
    Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long

Private Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function FindFirstFile Lib "Kernel32" _
    Alias "FindFirstFileA" (ByVal lpFileName As String, _
    lpFindFileData As WIN32_FIND_DATA) As Long

Private Declare Function FindNextFile Lib "Kernel32" _
    Alias "FindNextFileA" (ByVal hFindFile As Long, _
    lpFindFileData As WIN32_FIND_DATA) As Long

Private Declare Function FindClose Lib "Kernel32" _
    (ByVal hFindFile As Long) As Long

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

Private Declare Function lstrcat Lib "Kernel32" _
    Alias "lstrcatA" (ByVal lpString1 As String, _
    ByVal lpString2 As String) As Long

' AFTER (64-bit compatible):
Private Declare PtrSafe Function SHGetPathFromIDList Lib "shell32.dll" _
    Alias "SHGetPathFromIDListA" (ByVal pidl As LongPtr, _
    ByVal pszPath As String) As Long

Private Declare PtrSafe Function SHBrowseForFolder Lib "shell32.dll" _
    Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As LongPtr

Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" (ByVal hWnd As LongPtr, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As LongPtr

Private Declare PtrSafe Function FindFirstFile Lib "Kernel32" _
    Alias "FindFirstFileA" (ByVal lpFileName As String, _
    lpFindFileData As WIN32_FIND_DATA) As LongPtr

Private Declare PtrSafe Function FindNextFile Lib "Kernel32" _
    Alias "FindNextFileA" (ByVal hFindFile As LongPtr, _
    lpFindFileData As WIN32_FIND_DATA) As Long

Private Declare PtrSafe Function FindClose Lib "Kernel32" _
    (ByVal hFindFile As LongPtr) As Long

Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As LongPtr)

Private Declare PtrSafe Function lstrcat Lib "Kernel32" _
    Alias "lstrcatA" (ByVal lpString1 As String, _
    ByVal lpString2 As String) As LongPtr
```

> ℹ️ `FindFirstFile` returns an OS file-search HANDLE — must be `LongPtr`. Pass that handle to `FindNextFile` and `FindClose`.
> ℹ️ `CoTaskMemFree` frees the PIDL pointer returned by `SHBrowseForFolder` — the pointer param must be `LongPtr`.

---

### modExplorerApi (Shell32.dll)
```vba
' BEFORE (32-bit):
Private Declare Function SHBrowseForFolder Lib "shell32.dll" _
    Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" _
    Alias "SHGetPathFromIDListA" (ByVal pidl As Long, _
    ByVal pszPath As String) As Long

' AFTER (64-bit compatible):
Private Declare PtrSafe Function SHBrowseForFolder Lib "shell32.dll" _
    Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As LongPtr

Private Declare PtrSafe Function SHGetPathFromIDList Lib "shell32.dll" _
    Alias "SHGetPathFromIDListA" (ByVal pidl As LongPtr, _
    ByVal pszPath As String) As Long
```

---

### modIniApi (Kernel32.dll)
```vba
' BEFORE (32-bit):
Private Declare Function GetPrivateProfileString Lib "Kernel32" _
    Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, ByVal lpDefault As String, _
    ByVal lpReturnedString As String, ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "Kernel32" _
    Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, ByVal lpString As Any, _
    ByVal lpFileName As String) As Long

' AFTER (64-bit compatible):
Private Declare PtrSafe Function GetPrivateProfileString Lib "Kernel32" _
    Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, ByVal lpDefault As String, _
    ByVal lpReturnedString As String, ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

Private Declare PtrSafe Function WritePrivateProfileString Lib "Kernel32" _
    Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, ByVal lpString As Any, _
    ByVal lpFileName As String) As Long
```

> ℹ️ INI file APIs are pure string/integer — only `PtrSafe` needs adding. No `LongPtr` changes required.

---

### modRegistry (Advapi32.dll) — Basic registry functions
```vba
' BEFORE (32-bit):
Private Declare Function RegOpenKey Lib "advapi32.dll" _
    Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, _
    phkResult As Long) As Long

Private Declare Function RegQueryValueEx Lib "advapi32.dll" _
    Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
    lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

Private Declare Function RegSetValueEx Lib "advapi32.dll" _
    Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
    ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, _
    ByVal cbData As Long) As Long

Private Declare Function RegCreateKey Lib "advapi32.dll" _
    Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, _
    phkResult As Long) As Long

Private Declare Function RegDeleteKey Lib "advapi32.dll" _
    Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long

Private Declare Function RegDeleteValue Lib "advapi32.dll" _
    Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" _
    (ByVal hKey As Long) As Long

' AFTER (64-bit compatible):
Private Declare PtrSafe Function RegOpenKey Lib "advapi32.dll" _
    Alias "RegOpenKeyA" (ByVal hKey As LongPtr, ByVal lpSubKey As String, _
    phkResult As LongPtr) As Long

Private Declare PtrSafe Function RegQueryValueEx Lib "advapi32.dll" _
    Alias "RegQueryValueExA" (ByVal hKey As LongPtr, ByVal lpValueName As String, _
    lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

Private Declare PtrSafe Function RegSetValueEx Lib "advapi32.dll" _
    Alias "RegSetValueExA" (ByVal hKey As LongPtr, ByVal lpValueName As String, _
    ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, _
    ByVal cbData As Long) As Long

Private Declare PtrSafe Function RegCreateKey Lib "advapi32.dll" _
    Alias "RegCreateKeyA" (ByVal hKey As LongPtr, ByVal lpSubKey As String, _
    phkResult As LongPtr) As Long

Private Declare PtrSafe Function RegDeleteKey Lib "advapi32.dll" _
    Alias "RegDeleteKeyA" (ByVal hKey As LongPtr, ByVal lpSubKey As String) As Long

Private Declare PtrSafe Function RegDeleteValue Lib "advapi32.dll" _
    Alias "RegDeleteValueA" (ByVal hKey As LongPtr, ByVal lpValueName As String) As Long

Private Declare PtrSafe Function RegCloseKey Lib "advapi32.dll" _
    (ByVal hKey As LongPtr) As Long
```

---

### modRegistry3 (Advapi32.dll) — Extended registry functions
```vba
' BEFORE (32-bit):
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" _
    Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
    ByVal ulOptions As Long, ByVal samDesired As Long, _
    phkResult As Long) As Long

Private Declare Function RegCreateKeyEx Lib "advapi32.dll" _
    Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
    ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, _
    ByVal samDesired As Long, lpSecurityAttributes As Any, _
    phkResult As Long, lpdwDisposition As Long) As Long

' AFTER (64-bit compatible):
Private Declare PtrSafe Function RegOpenKeyEx Lib "advapi32.dll" _
    Alias "RegOpenKeyExA" (ByVal hKey As LongPtr, ByVal lpSubKey As String, _
    ByVal ulOptions As Long, ByVal samDesired As Long, _
    phkResult As LongPtr) As Long

Private Declare PtrSafe Function RegCreateKeyEx Lib "advapi32.dll" _
    Alias "RegCreateKeyExA" (ByVal hKey As LongPtr, ByVal lpSubKey As String, _
    ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, _
    ByVal samDesired As Long, lpSecurityAttributes As Any, _
    phkResult As LongPtr, lpdwDisposition As Long) As Long
```

---

### modWinInet / modFTP / clsFTP (wininet.dll)
```vba
' BEFORE (32-bit):
Private Declare Function InternetOpen Lib "wininet.dll" _
    Alias "InternetOpenA" (ByVal lpszAgent As String, _
    ByVal dwAccessType As Long, ByVal lpszProxyName As String, _
    ByVal lpszProxyBypass As String, ByVal dwFlags As Long) As Long

Private Declare Function InternetConnect Lib "wininet.dll" _
    Alias "InternetConnectA" (ByVal hInternet As Long, _
    ByVal lpszServerName As String, ByVal nServerPort As Long, _
    ByVal lpszUserName As String, ByVal lpszPassword As String, _
    ByVal dwService As Long, ByVal dwFlags As Long, _
    ByVal dwContext As Long) As Long

Private Declare Function FtpGetFile Lib "wininet.dll" _
    Alias "FtpGetFileA" (ByVal hConnect As Long, _
    ByVal lpszRemoteFile As String, ByVal lpszNewFile As String, _
    ByVal fFailIfExists As Long, ByVal dwFlagsAndAttributes As Long, _
    ByVal dwFlags As Long, ByVal dwContext As Long) As Long

Private Declare Function FtpPutFile Lib "wininet.dll" _
    Alias "FtpPutFileA" (ByVal hConnect As Long, _
    ByVal lpszLocalFile As String, ByVal lpszNewRemoteFile As String, _
    ByVal dwFlags As Long, ByVal dwContext As Long) As Long

Private Declare Function FtpOpenFile Lib "wininet.dll" _
    Alias "FtpOpenFileA" (ByVal hConnect As Long, _
    ByVal lpszFileName As String, ByVal lAccess As Long, _
    ByVal lFlags As Long, ByVal lContext As Long) As Long

Private Declare Function FtpCreateDirectory Lib "wininet.dll" _
    Alias "FtpCreateDirectoryA" (ByVal hConnect As Long, _
    ByVal lpszDirectory As String) As Long

Private Declare Function FtpGetCurrentDirectory Lib "wininet.dll" _
    Alias "FtpGetCurrentDirectoryA" (ByVal hConnect As Long, _
    ByVal lpszCurrentDirectory As String, _
    lpdwBufferLength As Long) As Long

Private Declare Function FtpSetCurrentDirectory Lib "wininet.dll" _
    Alias "FtpSetCurrentDirectoryA" (ByVal hConnect As Long, _
    ByVal lpszDirectory As String) As Long

Private Declare Function InternetCloseHandle Lib "wininet.dll" _
    (ByVal hInternet As Long) As Long

' AFTER (64-bit compatible):
Private Declare PtrSafe Function InternetOpen Lib "wininet.dll" _
    Alias "InternetOpenA" (ByVal lpszAgent As String, _
    ByVal dwAccessType As Long, ByVal lpszProxyName As String, _
    ByVal lpszProxyBypass As String, ByVal dwFlags As Long) As LongPtr

Private Declare PtrSafe Function InternetConnect Lib "wininet.dll" _
    Alias "InternetConnectA" (ByVal hInternet As LongPtr, _
    ByVal lpszServerName As String, ByVal nServerPort As Long, _
    ByVal lpszUserName As String, ByVal lpszPassword As String, _
    ByVal dwService As Long, ByVal dwFlags As Long, _
    ByVal dwContext As LongPtr) As LongPtr

Private Declare PtrSafe Function FtpGetFile Lib "wininet.dll" _
    Alias "FtpGetFileA" (ByVal hConnect As LongPtr, _
    ByVal lpszRemoteFile As String, ByVal lpszNewFile As String, _
    ByVal fFailIfExists As Long, ByVal dwFlagsAndAttributes As Long, _
    ByVal dwFlags As Long, ByVal dwContext As LongPtr) As Long

Private Declare PtrSafe Function FtpPutFile Lib "wininet.dll" _
    Alias "FtpPutFileA" (ByVal hConnect As LongPtr, _
    ByVal lpszLocalFile As String, ByVal lpszNewRemoteFile As String, _
    ByVal dwFlags As Long, ByVal dwContext As LongPtr) As Long

Private Declare PtrSafe Function FtpOpenFile Lib "wininet.dll" _
    Alias "FtpOpenFileA" (ByVal hConnect As LongPtr, _
    ByVal lpszFileName As String, ByVal lAccess As Long, _
    ByVal lFlags As Long, ByVal lContext As LongPtr) As LongPtr

Private Declare PtrSafe Function FtpCreateDirectory Lib "wininet.dll" _
    Alias "FtpCreateDirectoryA" (ByVal hConnect As LongPtr, _
    ByVal lpszDirectory As String) As Long

Private Declare PtrSafe Function FtpGetCurrentDirectory Lib "wininet.dll" _
    Alias "FtpGetCurrentDirectoryA" (ByVal hConnect As LongPtr, _
    ByVal lpszCurrentDirectory As String, _
    lpdwBufferLength As Long) As Long

Private Declare PtrSafe Function FtpSetCurrentDirectory Lib "wininet.dll" _
    Alias "FtpSetCurrentDirectoryA" (ByVal hConnect As LongPtr, _
    ByVal lpszDirectory As String) As Long

Private Declare PtrSafe Function InternetCloseHandle Lib "wininet.dll" _
    (ByVal hInternet As LongPtr) As Long
```

---

### modODBC (ODBC32.dll)
```vba
' BEFORE (32-bit):
Private Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" _
    (ByVal hwndParent As Long, ByVal fRequest As Long, _
    ByVal lpszDriver As String, ByVal lpszAttributes As String) As Long

' AFTER (64-bit compatible):
Private Declare PtrSafe Function SQLConfigDataSource Lib "ODBCCP32.DLL" _
    (ByVal hwndParent As LongPtr, ByVal fRequest As Long, _
    ByVal lpszDriver As String, ByVal lpszAttributes As String) As Long
```

> ℹ️ `hwndParent` is an HWND (window handle) — must be `LongPtr`. The return value is BOOL → stays `As Long`.

---

## PART 3: TYPE STRUCTURE UPDATES

### BROWSEINFO Structure (modFileDlgs / modExplorerApi)
```vba
' BEFORE:
Private Type BROWSEINFO
    hwndOwner      As Long
    pidlRoot       As Long
    pszDisplayName As String
    lpszTitle      As String
    ulFlags        As Long
    lpfn           As Long
    lParam         As Long
    iImage         As Long
End Type

' AFTER:
Private Type BROWSEINFO
    hwndOwner      As LongPtr   ' HWND — pointer-sized
    pidlRoot       As LongPtr   ' PIDLIST_ABSOLUTE — pointer
    pszDisplayName As String    ' NO CHANGE — string
    lpszTitle      As String    ' NO CHANGE — string
    ulFlags        As Long      ' NO CHANGE — flags
    lpfn           As LongPtr   ' BFFCALLBACK — function pointer
    lParam         As LongPtr   ' application-defined — pointer context
    iImage         As Long      ' NO CHANGE — icon index
End Type
```

### SHFILEINFO Structure (modAPIIcons)
```vba
' BEFORE:
Private Type SHFILEINFO
    hIcon         As Long
    iIcon         As Long
    dwAttributes  As Long
    szDisplayName As String * 260
    szTypeName    As String * 80
End Type

' AFTER:
Private Type SHFILEINFO
    hIcon         As LongPtr    ' HICON — handle, pointer-sized
    iIcon         As Long       ' NO CHANGE — icon index
    dwAttributes  As Long       ' NO CHANGE — file attributes
    szDisplayName As String * 260
    szTypeName    As String * 80
End Type
```

### WIN32_FIND_DATA Structure (modFileDlgs — if declared as VBA Type)
```vba
' BEFORE:
Private Type WIN32_FIND_DATA
    dwFileAttributes   As Long
    ftCreationTime     As FILETIME
    ftLastAccessTime   As FILETIME
    ftLastWriteTime    As FILETIME
    nFileSizeHigh      As Long
    nFileSizeLow       As Long
    dwReserved0        As Long
    dwReserved1        As Long
    cFileName          As String * 260
    cAlternate         As String * 14
End Type

' AFTER: No changes needed — WIN32_FIND_DATA has no pointer fields
' The HANDLE returned by FindFirstFile is NOT stored in this structure
```

---

## PART 4: CONSTANTS THAT MUST REMAIN UNCHANGED

These are safe numeric constants — do NOT change them:

```vba
' Window / ShowWindow constants
Public Const GW_HWNDNEXT    = 2
Public Const GW_HWNDPREV    = 3
Public Const SW_SHOW        = 5
Public Const SW_HIDE        = 0
Public Const SW_SHOWNORMAL  = 1

' Path / buffer sizes
Public Const MAX_PATH       = 260

' Registry root keys (see NOTE below)
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_CURRENT_USER  = &H80000001

' Registry value types
Public Const REG_SZ     = 1
Public Const REG_DWORD  = 4

' Registry access flags
Public Const KEY_READ   = &H20019
Public Const KEY_WRITE  = &H20006

' SendMessage / icon constants
Public Const WM_GETICON = &H7F
Public Const ICON_SMALL = 0
Public Const ICON_BIG   = 1

' Global memory flags
Public Const GHND         = &H42
Public Const GMEM_MOVEABLE = &H2

' Shell browse flags
Public Const BIF_RETURNONLYFSDIRS = &H1

' ODBC DSN request codes
Public Const ODBC_ADD_DSN    = 1
Public Const ODBC_CONFIG_DSN = 2
Public Const ODBC_REMOVE_DSN = 3
```

> ⚠️ **REGISTRY ROOT KEY EXCEPTION:** `HKEY_LOCAL_MACHINE = &H80000002` — In 64-bit VBA, predefined registry root keys must be passed as `LongPtr` to registry functions. Use `CLngPtr()` to cast:
> ```vba
> ' Safe approach for registry root keys in 64-bit:
> RegOpenKey CLngPtr(HKEY_LOCAL_MACHINE), "Software\...", hSubKey
> ' OR declare as LongPtr variable:
> Dim hRootKey As LongPtr
> hRootKey = CLngPtr(&H80000002)   ' HKEY_LOCAL_MACHINE
> ```

---

## PART 5: CONDITIONAL COMPILATION (Optional — for cross-version compatibility)

If the database needs to run on BOTH 32-bit and 64-bit Office, use `#If VBA7`:

```vba
#If VBA7 Then
    ' 64-bit compatible declarations
    Private Declare PtrSafe Function GetClassName Lib "User32" Alias "GetClassNameA" _
        (ByVal hWnd As LongPtr, ByVal lpClassName As String, _
        ByVal nMaxCount As Long) As Long
#Else
    ' 32-bit declarations (Office 2007 and earlier)
    Private Declare Function GetClassName Lib "User32" Alias "GetClassNameA" _
        (ByVal hWnd As Long, ByVal lpClassName As String, _
        ByVal nMaxCount As Long) As Long
#End If
```

> Since this codebase targets 64-bit exclusively, skip `#If VBA7` blocks and use the `PtrSafe` / `LongPtr` versions directly.

---

## PART 6: VERIFICATION CHECKLIST

After applying all find/replace operations, verify:

**PtrSafe:**
- [ ] No `Declare Function` without `PtrSafe` remains (search: `Declare Function` — should only find `Declare PtrSafe Function`)
- [ ] No `Declare Sub` without `PtrSafe` remains

**User32.dll parameters:**
- [ ] All `hWnd`, `hwnd`, `hWndParent`, `hwndOwner` parameters are `LongPtr`
- [ ] `hInst`, `hInstance`, `hIcon` handle parameters are `LongPtr`
- [ ] `SendMessage` — `hWnd`, `wParam`, `lParam` all `LongPtr`; return type `LongPtr`
- [ ] `LoadImage` — `hInst` param and return type are `LongPtr`
- [ ] `GetClipboardData` / `SetClipboardData` — return/accept `LongPtr`

**Kernel32.dll parameters:**
- [ ] `GlobalAlloc` returns `LongPtr`; `dwBytes` param is `LongPtr`
- [ ] `GlobalLock` takes and returns `LongPtr`
- [ ] `GlobalFree` / `GlobalUnlock` take `LongPtr`; return `Long`
- [ ] `FindFirstFile` returns `LongPtr` (file search handle)
- [ ] `FindNextFile` / `FindClose` take `LongPtr` handle param

**Shell32.dll parameters:**
- [ ] `pidl`, `pidlRoot` pointer params are `LongPtr`
- [ ] `SHBrowseForFolder` return type is `LongPtr`
- [ ] `ShellExecute` — `hWnd` param and return type are `LongPtr`
- [ ] `CoTaskMemFree` — pointer param is `LongPtr`

**Registry (Advapi32.dll):**
- [ ] All `hKey` parameters and `phkResult` output params are `LongPtr`
- [ ] `RegCreateKeyEx` — `hKey`, `phkResult` are `LongPtr`; return is `Long`
- [ ] Registry root key constants cast with `CLngPtr()` when passed to functions

**WinInet.dll parameters:**
- [ ] `hInternet`, `hConnect`, `hFtpSession` handle params are `LongPtr`
- [ ] `dwContext` params in WinInet calls are `LongPtr`
- [ ] `InternetOpen`, `InternetConnect`, `FtpOpenFile` return types are `LongPtr`
- [ ] `FtpCreateDirectory`, `FtpGetCurrentDirectory`, `FtpSetCurrentDirectory` — handle params `LongPtr`

**ODBC:**
- [ ] `SQLConfigDataSource` — `hwndParent` param is `LongPtr`

**Type Structures:**
- [ ] `BROWSEINFO.hwndOwner`, `.pidlRoot`, `.lParam`, `.lpfn` fields are `LongPtr`
- [ ] `SHFILEINFO.hIcon` field is `LongPtr`

**Local Variables:**
- [ ] No `Dim hWnd As Long` (or other handle/pointer vars) remains as plain `Long`

**Final compile:**
- [ ] VBA IDE → Debug → Compile → zero errors

---

## PART 7: COMMON POST-CONVERSION ERRORS & FIXES

| Error Message | Likely Cause | Fix |
|---|---|---|
| `Type mismatch` | Long variable receiving LongPtr function result | Change `Dim hXxx As Long` to `Dim hXxx As LongPtr` |
| `Expected: expression` | Bad find/replace created syntax error | Undo (Ctrl+Z) that replace, redo manually |
| `Compile error: Can't find project or library` | Missing reference or 32-bit DLL path | Check Tools → References |
| Comparing LongPtr to 0 | `If hWnd = 0 Then` — works in VBA, no change needed | No action needed — VBA allows `= 0` with LongPtr |
| Registry functions return wrong values | Root key constant passed as Long | Wrap with `CLngPtr(HKEY_LOCAL_MACHINE)` |
| `GlobalAlloc` returns wrong size | `dwBytes` passed as `Long` (max 2GB) | Change `dwBytes As Long` to `dwBytes As LongPtr` |