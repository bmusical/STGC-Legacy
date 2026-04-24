# VBA 32-bit to 64-bit API Conversion Guide
## STGC-Legacy (ConDispatchSQL) - Access VBA Migration

---

## Overview

This document provides a complete set of **find/replace operations** to convert all 32-bit Windows API `Declare` statements to 64-bit compatible equivalents across the following modules identified in this codebase:

| Module | APIs Used |
|---|---|
| `modAPI` | User32 (GetClassNameA, GetWindow, ShowWindow, etc.) |
| `modAPIIcons` | User32, Shell32 (LoadImage, SHGetFileInfo) |
| `modClipboard` | User32 (OpenClipboard, CloseClipboard, GetClipboardData, etc.) |
| `modExplorerApi` | Shell32 (ShellExecute, SHGetPathFromIDList) |
| `modFileDlgs` | Shell32 (SHGetPathFromIDList, SHBrowseForFolder) |
| `modIniApi` | Kernel32 (GetPrivateProfileString, WritePrivateProfileString) |
| `modRegistry` / `modRegistry3` | Advapi32 (RegOpenKey, RegQueryValueEx, RegSetValueEx, etc.) |
| `modWinInet` / `modFTP` | WinInet (InternetOpen, FtpGetFile, etc.) |

---

## THE GOLDEN RULE OF 32→64 BIT VBA CONVERSION

The core changes required are:

1. Add **`PtrSafe`** after `Declare` (or confirm it's already there)
2. Change **`As Long`** to **`As LongPtr`** for all **handle**, **pointer**, and **HWND** parameters
3. Change **`ByVal hWnd As Long`** → **`ByVal hWnd As LongPtr`** everywhere
4. Keep **`As Long`** for pure integer/numeric values (counts, flags, error codes)
5. Change **`As Long`** to **`As LongPtr`** for function return types that return handles/pointers

---

## PART 1: FIND/REPLACE OPERATIONS (Apply in This Exact Order)

### STEP 1 — Add PtrSafe to all Declare statements missing it
> Run this FIRST. Any `Declare` without `PtrSafe` must get it added.

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

> ⚠️ Skip if already has `PtrSafe` — use "whole word" / case-insensitive match.
> After this step, verify no `Declare Function` or `Declare Sub` remains without `PtrSafe`.

---

### STEP 2 — Fix HWND parameters (most common handle type)
> `HWND` is a pointer-sized value — must be `LongPtr` in 64-bit.

**Find:**
```
ByVal hWnd As Long
```
**Replace With:**
```
ByVal hWnd As LongPtr
```

**Find:**
```
ByVal hwnd As Long
```
**Replace With:**
```
ByVal hwnd As LongPtr
```

**Find:**
```
ByVal hWndParent As Long
```
**Replace With:**
```
ByVal hWndParent As LongPtr
```

**Find:**
```
ByVal hwndOwner As Long
```
**Replace With:**
```
ByVal hwndOwner As LongPtr
```

---

### STEP 3 — Fix HANDLE parameters
> All Windows HANDLE types (HINSTANCE, HICON, HBITMAP, HMENU, etc.) are pointer-sized.

**Find:**
```
ByVal hInst As Long
```
**Replace With:**
```
ByVal hInst As LongPtr
```

**Find:**
```
ByVal hInstance As Long
```
**Replace With:**
```
ByVal hInstance As LongPtr
```

**Find:**
```
ByVal hIcon As Long
```
**Replace With:**
```
ByVal hIcon As LongPtr
```

**Find:**
```
ByVal hBitmap As Long
```
**Replace With:**
```
ByVal hBitmap As LongPtr
```

**Find:**
```
ByVal hMenu As Long
```
**Replace With:**
```
ByVal hMenu As LongPtr
```

**Find:**
```
ByVal hModule As Long
```
**Replace With:**
```
ByVal hModule As LongPtr
```

**Find:**
```
ByVal hKey As Long
```
**Replace With:**
```
ByVal hKey As LongPtr
```

**Find:**
```
ByRef hKey As Long
```
**Replace With:**
```
ByRef hKey As LongPtr
```

**Find:**
```
ByVal hConn As Long
```
**Replace With:**
```
ByVal hConn As LongPtr
```

**Find:**
```
ByVal hInternet As Long
```
**Replace With:**
```
ByVal hInternet As LongPtr
```

**Find:**
```
ByVal hFtpSession As Long
```
**Replace With:**
```
ByVal hFtpSession As LongPtr
```

---

### STEP 4 — Fix pointer/buffer parameters (LPVOID, LPCSTR used as Long)

**Find:**
```
ByVal lpszAgent As Long
```
**Replace With:**
```
ByVal lpszAgent As LongPtr
```

**Find:**
```
ByVal lParam As Long
```
**Replace With:**
```
ByVal lParam As LongPtr
```

**Find:**
```
ByVal wParam As Long
```
**Replace With:**
```
ByVal wParam As LongPtr
```

**Find:**
```
ByVal pidl As Long
```
**Replace With:**
```
ByVal pidl As LongPtr
```

**Find:**
```
ByRef pidl As Long
```
**Replace With:**
```
ByRef pidl As LongPtr
```

---

### STEP 5 — Fix function return types that return handles/pointers

**Find:**
```
) As Long  'GetWindow
```
> This requires manual review. For all `Declare Function` lines, if the return type `As Long` represents a handle or pointer, change it to `As LongPtr`.

Common functions whose return type must be `LongPtr`:
- `GetWindow`, `FindWindow`, `FindWindowEx` → return HWND → `As LongPtr`
- `ShellExecute`, `ShellExecuteEx` → return HINSTANCE → `As LongPtr`  
- `LoadImage`, `LoadIcon`, `LoadBitmap` → return HANDLE → `As LongPtr`
- `SHGetFileInfo` → return DWORD_PTR → `As LongPtr`
- `InternetOpen`, `InternetConnect`, `FtpOpenFile` → return HINTERNET → `As LongPtr`
- `RegOpenKey`, `RegOpenKeyEx` → return LONG (error code) → keep `As Long`
- `GetClipboardData` → return HANDLE → `As LongPtr`
- `OpenClipboard` → return BOOL (integer) → keep `As Long`

---

### STEP 6 — Fix Type structures containing handles (UDTs)

For any `Type` / `End Type` blocks containing handle fields:

**Find (in SHFILEINFO type):**
```
hIcon As Long
```
**Replace With:**
```
hIcon As LongPtr
```

**Find (in BROWSEINFO type):**
```
hwndOwner As Long
```
**Replace With:**
```
hwndOwner As LongPtr
```

**Find (in BROWSEINFO type):**
```
pidlRoot As Long
```
**Replace With:**
```
pidlRoot As LongPtr
```

**Find (in BROWSEINFO type):**
```
lParam As Long
```
**Replace With:**
```
lParam As LongPtr
```

---

### STEP 7 — Fix variable declarations in code (not just Declare lines)
> Any variable that holds a value returned from an API that now returns `LongPtr`.

**Find:**
```
Dim hWnd As Long
```
**Replace With:**
```
Dim hWnd As LongPtr
```

**Find:**
```
Dim hwnd As Long
```
**Replace With:**
```
Dim hwnd As LongPtr
```

**Find:**
```
Dim hInst As Long
```
**Replace With:**
```
Dim hInst As LongPtr
```

**Find:**
```
Dim hIcon As Long
```
**Replace With:**
```
Dim hIcon As LongPtr
```

**Find:**
```
Dim hInternet As Long
```
**Replace With:**
```
Dim hInternet As LongPtr
```

**Find:**
```
Dim hFTP As Long
```
**Replace With:**
```
Dim hFTP As LongPtr
```

**Find:**
```
Dim hKey As Long
```
**Replace With:**
```
Dim hKey As LongPtr
```

---

## PART 2: MODULE-BY-MODULE REFERENCE

### modAPI (User32.dll)
```vba
' BEFORE (32-bit):
Private Declare Function GetClassName Lib "User32" (ByVal hWnd As Long, _
    ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Private Declare Function GetWindow Lib "User32" (ByVal hWnd As Long, _
    ByVal wCmd As Long) As Long

Private Declare Function ShowWindow Lib "User32" (ByVal hWnd As Long, _
    ByVal nCmdShow As Long) As Boolean

' AFTER (64-bit compatible):
Private Declare PtrSafe Function GetClassName Lib "User32" Alias "GetClassNameA" _
    (ByVal hWnd As LongPtr, ByVal lpClassName As String, _
    ByVal nMaxCount As Long) As Long

Private Declare PtrSafe Function GetWindow Lib "User32" (ByVal hWnd As LongPtr, _
    ByVal wCmd As Long) As LongPtr

Private Declare PtrSafe Function ShowWindow Lib "User32" (ByVal hWnd As LongPtr, _
    ByVal nCmdShow As Long) As Long
```

---

### modClipboard (User32.dll)
```vba
' BEFORE (32-bit):
Private Declare Function OpenClipboard Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function CloseClipboard Lib "User32" () As Long
Private Declare Function GetClipboardData Lib "User32" (ByVal wFormat As Long) As Long
Private Declare Function SetClipboardData Lib "User32" (ByVal wFormat As Long, _
    ByVal hMem As Long) As Long
Private Declare Function EmptyClipboard Lib "User32" () As Long

' AFTER (64-bit compatible):
Private Declare PtrSafe Function OpenClipboard Lib "User32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "User32" () As Long
Private Declare PtrSafe Function GetClipboardData Lib "User32" (ByVal wFormat As Long) As LongPtr
Private Declare PtrSafe Function SetClipboardData Lib "User32" (ByVal wFormat As Long, _
    ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function EmptyClipboard Lib "User32" () As Long
```

---

### modFileDlgs / modExplorerApi (Shell32.dll)
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
> ℹ️ INI file APIs are pure string/integer — no LongPtr needed except adding PtrSafe.

---

### modRegistry / modRegistry3 (Advapi32.dll)
```vba
' BEFORE (32-bit):
Private Declare Function RegOpenKey Lib "advapi32.dll" _
    Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, _
    phkResult As Long) As Long

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" _
    Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
    ByVal ulOptions As Long, ByVal samDesired As Long, _
    phkResult As Long) As Long

Private Declare Function RegQueryValueEx Lib "advapi32.dll" _
    Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
    lpReserved As Long, lpType As Long, lpData As Any, _
    lpcbData As Long) As Long

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

Private Declare PtrSafe Function RegOpenKeyEx Lib "advapi32.dll" _
    Alias "RegOpenKeyExA" (ByVal hKey As LongPtr, ByVal lpSubKey As String, _
    ByVal ulOptions As Long, ByVal samDesired As Long, _
    phkResult As LongPtr) As Long

Private Declare PtrSafe Function RegQueryValueEx Lib "advapi32.dll" _
    Alias "RegQueryValueExA" (ByVal hKey As LongPtr, ByVal lpValueName As String, _
    lpReserved As Long, lpType As Long, lpData As Any, _
    lpcbData As Long) As Long

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

### modWinInet / modFTP (wininet.dll)
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

Private Declare PtrSafe Function InternetCloseHandle Lib "wininet.dll" _
    (ByVal hInternet As LongPtr) As Long
```

---

### modAPIIcons (User32.dll / Shell32.dll)
```vba
' BEFORE (32-bit):
Private Declare Function LoadImage Lib "User32" Alias "LoadImageA" _
    (ByVal hInst As Long, ByVal lpszName As String, _
    ByVal uType As Long, ByVal cxDesired As Long, _
    ByVal cyDesired As Long, ByVal fuLoad As Long) As Long

Private Declare Function SHGetFileInfo Lib "shell32.dll" _
    Alias "SHGetFileInfoA" (ByVal pszPath As String, _
    ByVal dwFileAttributes As Long, psfi As SHFILEINFO, _
    ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long

Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" _
    (ByVal hWnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, lParam As Long) As Long

' AFTER (64-bit compatible):
Private Declare PtrSafe Function LoadImage Lib "User32" Alias "LoadImageA" _
    (ByVal hInst As LongPtr, ByVal lpszName As String, _
    ByVal uType As Long, ByVal cxDesired As Long, _
    ByVal cyDesired As Long, ByVal fuLoad As Long) As LongPtr

Private Declare PtrSafe Function SHGetFileInfo Lib "shell32.dll" _
    Alias "SHGetFileInfoA" (ByVal pszPath As String, _
    ByVal dwFileAttributes As Long, psfi As SHFILEINFO, _
    ByVal cbFileInfo As Long, ByVal uFlags As Long) As LongPtr

Private Declare PtrSafe Function SendMessage Lib "User32" Alias "SendMessageA" _
    (ByVal hWnd As LongPtr, ByVal wMsg As Long, _
    ByVal wParam As LongPtr, lParam As LongPtr) As LongPtr
```

---

## PART 3: TYPE STRUCTURE UPDATES

### BROWSEINFO Structure (modFileDlgs)
```vba
' BEFORE:
Private Type BROWSEINFO
    hwndOwner     As Long
    pidlRoot      As Long
    pszDisplayName As String
    lpszTitle     As String
    ulFlags       As Long
    lpfn          As Long
    lParam        As Long
    iImage        As Long
End Type

' AFTER:
Private Type BROWSEINFO
    hwndOwner     As LongPtr
    pidlRoot      As LongPtr
    pszDisplayName As String
    lpszTitle     As String
    ulFlags       As Long
    lpfn          As LongPtr
    lParam        As LongPtr
    iImage        As Long
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
    hIcon         As LongPtr
    iIcon         As Long
    dwAttributes  As Long
    szDisplayName As String * 260
    szTypeName    As String * 80
End Type
```

---

## PART 4: CONSTANTS THAT MUST REMAIN UNCHANGED

These are safe and do NOT need modification:
```vba
' These stay As Long or As Integer - they are pure numeric constants/flags:
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const SW_SHOW = 5
Public Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1
Public Const MAX_PATH = 260
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_CURRENT_USER = &H80000001
Public Const REG_SZ = 1
Public Const REG_DWORD = 4
Public Const KEY_READ = &H20019
Public Const KEY_WRITE = &H20006
Public Const WM_GETICON = &H7F
Public Const ICON_SMALL = 0
Public Const ICON_BIG = 1
```

> ⚠️ **EXCEPTION:** `HKEY_LOCAL_MACHINE = &H80000002` — In 64-bit VBA, predefined registry root keys like `HKEY_LOCAL_MACHINE` and `HKEY_CURRENT_USER` need to be declared as `LongPtr` **variables** (not `Const`) when passed to registry functions, OR cast with `CLngPtr()`:
```vba
' Safe approach for registry root keys in 64-bit:
Dim hRootKey As LongPtr
hRootKey = CLngPtr(&H80000002)  ' HKEY_LOCAL_MACHINE
```

---

## PART 5: CONDITIONAL COMPILATION (Optional — for cross-version compatibility)

If the database needs to run on BOTH 32-bit and 64-bit Office, use conditional compilation:

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

> Since you are targeting 64-bit exclusively, you can skip the `#If VBA7` blocks and just use the `PtrSafe` / `LongPtr` versions directly.

---

## PART 6: QUICK VERIFICATION CHECKLIST

After applying all find/replace operations, verify:

- [ ] No `Declare Function` without `PtrSafe` remains
- [ ] No `Declare Sub` without `PtrSafe` remains  
- [ ] All `hWnd`, `hwnd`, `hWndParent`, `hwndOwner` parameters are `LongPtr`
- [ ] All `hInst`, `hInstance`, `hIcon`, `hBitmap`, `hMenu` parameters are `LongPtr`
- [ ] All `hKey` parameters and `phkResult` output params are `LongPtr`
- [ ] All `hInternet`, `hConnect`, `hFTP` handle params are `LongPtr`
- [ ] All `pidl`, `pidlRoot` pointer params are `LongPtr`
- [ ] All `lParam`, `wParam` in `SendMessage` are `LongPtr`
- [ ] `GetClipboardData` and `SetClipboardData` return/accept `LongPtr`
- [ ] `BROWSEINFO.hwndOwner`, `.pidlRoot`, `.lParam` fields are `LongPtr`
- [ ] `SHFILEINFO.hIcon` field is `LongPtr`
- [ ] No `Dim hWnd As Long` (or other handle vars) remains as plain `Long`
- [ ] Registry root key constants cast with `CLngPtr()` when passed to functions
- [ ] Compile test: Open VBA IDE → Debug → Compile (no errors)

---

## PART 7: SUMMARY TABLE — What Changes and What Stays

| Parameter Type | 32-bit | 64-bit | Notes |
|---|---|---|---|
| Window handle (HWND) | `As Long` | `As LongPtr` | Always changes |
| Any HANDLE type | `As Long` | `As LongPtr` | Always changes |
| Pointer (LPVOID, etc.) | `As Long` | `As LongPtr` | Always changes |
| lParam / wParam | `As Long` | `As LongPtr` | Always changes |
| pidl (shell pointer) | `As Long` | `As LongPtr` | Always changes |
| Registry key (HKEY) | `As Long` | `As LongPtr` | Always changes |
| Internet handle | `As Long` | `As LongPtr` | Always changes |
| String length/count | `As Long` | `As Long` | NO CHANGE |
| Flags / bit masks | `As Long` | `As Long` | NO CHANGE |
| Boolean return (BOOL) | `As Long` | `As Long` | NO CHANGE |
| Error codes | `As Long` | `As Long` | NO CHANGE |
| Integer counts | `As Long` | `As Long` | NO CHANGE |
| `Declare Function` | (no PtrSafe) | `PtrSafe` added | Always changes |
| `Declare Sub` | (no PtrSafe) | `PtrSafe` added | Always changes |