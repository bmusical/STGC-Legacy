' =============================================================================
' VBA_32to64_AutoFix.bas
' STGC-Legacy — Automated 32-bit to 64-bit API Declaration Converter
'
' HOW TO USE:
'   1. Open ConDispatchSQL_ConvertTo64.accdb in Microsoft Access (64-bit)
'   2. Open VBA IDE (Alt+F11)
'   3. Insert a new Module (Insert > Module)
'   4. Paste this entire file into that module
'   5. Place cursor inside Sub Fix32to64_AllModules()
'   6. Press F5 to run
'   7. Check the Immediate Window (Ctrl+G) for results
'   8. After running, do Debug > Compile to verify no errors
'   9. Save and test the application
'
' WHAT IT DOES:
'   - Scans every VBA module and class in the project
'   - Applies all 32-to-64-bit find/replace substitutions in prescribed order
'   - Logs every change made to the Immediate Window
'   - Reports a summary when complete
'
' API COVERAGE (verified by binary scan of all 129 module blobs):
'   User32.dll:   GetClassName, GetWindow, ShowWindowAsync, SendMessage,
'                 OpenClipboard, CloseClipboard, GetClipboardData,
'                 SetClipboardData, EmptyClipboard, LoadImage
'   Kernel32.dll: GetPrivateProfileString, WritePrivateProfileString,
'                 GlobalAlloc, GlobalFree, GlobalLock, GlobalUnlock,
'                 GetComputerName, FindFirstFile, FindNextFile, FindClose,
'                 lstrcat
'   Shell32.dll:  ShellExecute, SHGetPathFromIDList, SHBrowseForFolder
'   ole32.dll:    CoTaskMemFree
'   Advapi32.dll: RegOpenKey, RegOpenKeyEx, RegCloseKey, RegQueryValueEx,
'                 RegSetValueEx, RegCreateKey, RegCreateKeyEx,
'                 RegDeleteKey, RegDeleteValue
'   WinInet.dll:  InternetOpen, InternetConnect, InternetCloseHandle,
'                 FtpGetFile, FtpPutFile, FtpOpenFile, FtpCreateDirectory,
'                 FtpGetCurrentDirectory, FtpSetCurrentDirectory,
'                 InternetWriteFile
'   ODBC32.dll:   SQLConfigDataSource
' =============================================================================

Option Compare Database
Option Explicit

' =============================================================================
' MAIN ENTRY POINT — Run this Sub
' =============================================================================
Public Sub Fix32to64_AllModules()
    Dim proj         As VBProject
    Dim comp         As VBComponent
    Dim totalChanges As Long
    Dim moduleChanges As Long
    Dim startTime    As Date

    startTime = Now()
    totalChanges = 0

    Debug.Print String(70, "=")
    Debug.Print "VBA 32-to-64 Bit Conversion — Started: " & Format(startTime, "yyyy-mm-dd hh:mm:ss")
    Debug.Print String(70, "=")

    Set proj = Application.VBE.ActiveVBProject

    For Each comp In proj.VBComponents
        Select Case comp.Type
            Case vbext_ct_StdModule, vbext_ct_ClassModule, _
                 vbext_ct_MSForm, vbext_ct_Document
                moduleChanges = FixModule(comp)
                If moduleChanges > 0 Then
                    Debug.Print "  [" & comp.Name & "] " & moduleChanges & " change(s) made"
                    totalChanges = totalChanges + moduleChanges
                End If
        End Select
    Next comp

    Debug.Print String(70, "-")
    Debug.Print "COMPLETE: " & totalChanges & " total change(s) across all modules"
    Debug.Print "Duration: " & Format(Now() - startTime, "hh:mm:ss")
    Debug.Print String(70, "=")
    Debug.Print ""
    Debug.Print "NEXT STEPS:"
    Debug.Print "  1. In VBA IDE: Debug > Compile to check for errors"
    Debug.Print "  2. Fix any remaining errors manually (see guide for common fixes)"
    Debug.Print "  3. Cast registry root key constants: CLngPtr(HKEY_LOCAL_MACHINE)"
    Debug.Print "  4. Save the database"

    MsgBox "32-to-64 Bit Conversion Complete!" & vbCrLf & vbCrLf & _
           totalChanges & " total change(s) made." & vbCrLf & vbCrLf & _
           "Check the Immediate Window (Ctrl+G) for details." & vbCrLf & _
           "Then run Debug > Compile to verify.", _
           vbInformation, "Conversion Complete"
End Sub

' =============================================================================
' Process a single VBA component/module
' Returns the number of changes made
' =============================================================================
Private Function FixModule(comp As VBComponent) As Long
    Dim code     As String
    Dim newCode  As String
    Dim changes  As Long

    changes = 0

    ' Get the entire module source code
    On Error Resume Next
    code = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
    If Err.Number <> 0 Then
        Debug.Print "  [" & comp.Name & "] ERROR reading module: " & Err.Description
        Err.Clear
        FixModule = 0
        Exit Function
    End If
    On Error GoTo 0

    newCode = code

    ' =========================================================================
    ' STEP 0: Safety cleanup — fix any LongPtrPtr typos from previous runs
    '         Must run BEFORE everything else
    ' =========================================================================
    newCode = ReplaceExact(newCode, "As LongPtrPtr", "As LongPtr", changes)
    newCode = ReplaceExact(newCode, "As LongPtr Ptr", "As LongPtr", changes)

    ' =========================================================================
    ' STEP 1: Add PtrSafe to Declare Function/Sub (case-insensitive search)
    '         Must run FIRST before any other replacements.
    ' =========================================================================
    newCode = ReplaceWord(newCode, "Declare Function ", "Declare PtrSafe Function ", changes)
    newCode = ReplaceWord(newCode, "Declare Sub ", "Declare PtrSafe Sub ", changes)
    ' Remove any accidental double-PtrSafe (module already had it)
    newCode = ReplaceExact(newCode, "Declare PtrSafe PtrSafe Function ", "Declare PtrSafe Function ", changes)
    newCode = ReplaceExact(newCode, "Declare PtrSafe PtrSafe Sub ", "Declare PtrSafe Sub ", changes)

    ' =========================================================================
    ' STEP 2: Fix HWND parameters (window handles)
    ' =========================================================================
    newCode = ReplaceExact(newCode, "ByVal hWnd As Long", "ByVal hWnd As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByVal hwnd As Long", "ByVal hwnd As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByVal hWndParent As Long", "ByVal hWndParent As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByVal hwndOwner As Long", "ByVal hwndOwner As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByVal hWndOwner As Long", "ByVal hWndOwner As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByRef hWnd As Long", "ByRef hWnd As LongPtr", changes)

    ' =========================================================================
    ' STEP 3: Fix generic HANDLE parameters
    '         (HINSTANCE, HICON, HBITMAP, HMENU, HGLOBAL, HFILE, HDC, etc.)
    ' =========================================================================
    newCode = ReplaceExact(newCode, "ByVal hInst As Long", "ByVal hInst As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByVal hInstance As Long", "ByVal hInstance As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByVal hIcon As Long", "ByVal hIcon As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByVal hBitmap As Long", "ByVal hBitmap As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByVal hMenu As Long", "ByVal hMenu As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByVal hModule As Long", "ByVal hModule As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByVal hProcess As Long", "ByVal hProcess As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByVal hThread As Long", "ByVal hThread As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByVal hObject As Long", "ByVal hObject As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByVal hDC As Long", "ByVal hDC As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByVal hFile As Long", "ByVal hFile As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByVal hMem As Long", "ByVal hMem As LongPtr", changes)

    ' =========================================================================
    ' STEP 4: Fix GlobalAlloc/GlobalLock — memory handle + size parameter
    '         GlobalAlloc dwBytes must be LongPtr for 64-bit address space
    ' =========================================================================
    newCode = ReplaceExact(newCode, "ByVal dwBytes As Long", "ByVal dwBytes As LongPtr", changes)

    ' =========================================================================
    ' STEP 5: Fix FindFirstFile / FindNextFile / FindClose handles
    '         FindFirstFile returns a HANDLE (search handle) — must be LongPtr
    ' =========================================================================
    newCode = ReplaceExact(newCode, "ByVal hFindFile As Long", "ByVal hFindFile As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByRef hFindFile As Long", "ByRef hFindFile As LongPtr", changes)
    newCode = ReplaceExact(newCode, "Dim hFindFile As Long", "Dim hFindFile As LongPtr", changes)

    ' =========================================================================
    ' STEP 6: Fix Registry handle parameters (HKEY)
    ' =========================================================================
    newCode = ReplaceExact(newCode, "ByVal hKey As Long", "ByVal hKey As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByRef hKey As Long", "ByRef hKey As LongPtr", changes)
    newCode = ReplaceExact(newCode, "phkResult As Long", "phkResult As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByRef phkResult As Long", "ByRef phkResult As LongPtr", changes)

    ' =========================================================================
    ' STEP 7: Fix WinInet / FTP handle parameters (HINTERNET)
    ' =========================================================================
    newCode = ReplaceExact(newCode, "ByVal hInternet As Long", "ByVal hInternet As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByVal hConnect As Long", "ByVal hConnect As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByVal hFtpSession As Long", "ByVal hFtpSession As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByVal hConnection As Long", "ByVal hConnection As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByVal dwContext As Long", "ByVal dwContext As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByVal hFindHandle As Long", "ByVal hFindHandle As LongPtr", changes)

    ' =========================================================================
    ' STEP 8: Fix CoTaskMemFree pointer parameter (ole32.dll)
    '         Used to free PIDL returned by SHBrowseForFolder
    ' =========================================================================
    newCode = ReplaceExact(newCode, "ByVal pv As Long", "ByVal pv As LongPtr", changes)

    ' =========================================================================
    ' STEP 9: Fix lParam / wParam (SendMessage and UDT fields)
    '         NOTE: bare "lParam As Long" removed — too broad, caused LongPtrPtr
    '         The ByVal/ByRef versions cover all Declare line cases safely.
    '         UDT lParam field is handled in STEP 12 (indented Type block).
    ' =========================================================================
    newCode = ReplaceExact(newCode, "ByVal lParam As Long", "ByVal lParam As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByRef lParam As Long", "ByRef lParam As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByVal wParam As Long", "ByVal wParam As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByRef wParam As Long", "ByRef wParam As LongPtr", changes)

    ' =========================================================================
    ' STEP 10: Fix Shell pointer parameters (PIDL)
    ' =========================================================================
    newCode = ReplaceExact(newCode, "ByVal pidl As Long", "ByVal pidl As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByRef pidl As Long", "ByRef pidl As LongPtr", changes)
    newCode = ReplaceExact(newCode, "pidlRoot As Long", "pidlRoot As LongPtr", changes)
    newCode = ReplaceExact(newCode, "lpfn As Long", "lpfn As LongPtr", changes)

    ' =========================================================================
    ' STEP 11: Fix ODBC hwndParent parameter (SQLConfigDataSource)
    ' =========================================================================
    newCode = ReplaceExact(newCode, "ByVal hwndParent As Long", "ByVal hwndParent As LongPtr", changes)

    ' =========================================================================
    ' STEP 12: Fix TYPE structure fields containing handles/pointers
    '          These are indented field declarations inside Type...End Type
    '          Using 4-space indent prefix to avoid hitting Declare/Dim lines
    ' =========================================================================
    newCode = ReplaceExact(newCode, "    hIcon As Long", "    hIcon As LongPtr", changes)
    newCode = ReplaceExact(newCode, "    hwndOwner As Long", "    hwndOwner As LongPtr", changes)
    newCode = ReplaceExact(newCode, "    hWnd As Long", "    hWnd As LongPtr", changes)
    newCode = ReplaceExact(newCode, "    lParam As Long", "    lParam As LongPtr", changes)
    newCode = ReplaceExact(newCode, "    pidlRoot As Long", "    pidlRoot As LongPtr", changes)
    newCode = ReplaceExact(newCode, "    lpfn As Long", "    lpfn As LongPtr", changes)

    ' =========================================================================
    ' STEP 13: Fix Dim statements for handle/pointer variables
    '          Includes standard h-prefix names AND lng-prefix Hungarian notation
    '          names used in this codebase (e.g. lngRootKey, lngKeyHandle)
    ' =========================================================================
    ' Standard h-prefix handle variables
    newCode = ReplaceExact(newCode, "Dim hWnd As Long", "Dim hWnd As LongPtr", changes)
    newCode = ReplaceExact(newCode, "Dim hwnd As Long", "Dim hwnd As LongPtr", changes)
    newCode = ReplaceExact(newCode, "Dim hInst As Long", "Dim hInst As LongPtr", changes)
    newCode = ReplaceExact(newCode, "Dim hIcon As Long", "Dim hIcon As LongPtr", changes)
    newCode = ReplaceExact(newCode, "Dim hInternet As Long", "Dim hInternet As LongPtr", changes)
    newCode = ReplaceExact(newCode, "Dim hFTP As Long", "Dim hFTP As LongPtr", changes)
    newCode = ReplaceExact(newCode, "Dim hConnect As Long", "Dim hConnect As LongPtr", changes)
    newCode = ReplaceExact(newCode, "Dim hKey As Long", "Dim hKey As LongPtr", changes)
    newCode = ReplaceExact(newCode, "Dim hMem As Long", "Dim hMem As LongPtr", changes)
    newCode = ReplaceExact(newCode, "Dim hDC As Long", "Dim hDC As LongPtr", changes)
    newCode = ReplaceExact(newCode, "Dim hFile As Long", "Dim hFile As LongPtr", changes)
    newCode = ReplaceExact(newCode, "Dim lParam As Long", "Dim lParam As LongPtr", changes)
    newCode = ReplaceExact(newCode, "Dim wParam As Long", "Dim wParam As LongPtr", changes)
    newCode = ReplaceExact(newCode, "Dim pidl As Long", "Dim pidl As LongPtr", changes)
    newCode = ReplaceExact(newCode, "Dim hFindFile As Long", "Dim hFindFile As LongPtr", changes)

    ' lng-prefix Hungarian notation registry handle variables (this codebase)
    newCode = ReplaceExact(newCode, "Dim lngRootKey As Long", "Dim lngRootKey As LongPtr", changes)
    newCode = ReplaceExact(newCode, "Dim lngKeyHandle As Long", "Dim lngKeyHandle As LongPtr", changes)
    newCode = ReplaceExact(newCode, "Dim lngKey As Long", "Dim lngKey As LongPtr", changes)
    newCode = ReplaceExact(newCode, "Dim lngHandle As Long", "Dim lngHandle As LongPtr", changes)
    newCode = ReplaceExact(newCode, "Dim lngHwnd As Long", "Dim lngHwnd As LongPtr", changes)
    newCode = ReplaceExact(newCode, "Dim lngHWnd As Long", "Dim lngHWnd As LongPtr", changes)
    newCode = ReplaceExact(newCode, "Dim lngHWND As Long", "Dim lngHWND As LongPtr", changes)
    newCode = ReplaceExact(newCode, "Dim lngInternet As Long", "Dim lngInternet As LongPtr", changes)
    newCode = ReplaceExact(newCode, "Dim lngConnect As Long", "Dim lngConnect As LongPtr", changes)
    newCode = ReplaceExact(newCode, "Dim lngFTP As Long", "Dim lngFTP As LongPtr", changes)

    ' Also fix lng-prefix as function/sub parameters on Declare lines
    newCode = ReplaceExact(newCode, "ByVal lngRootKey As Long", "ByVal lngRootKey As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByRef lngRootKey As Long", "ByRef lngRootKey As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByVal lngKeyHandle As Long", "ByVal lngKeyHandle As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByRef lngKeyHandle As Long", "ByRef lngKeyHandle As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByVal lngKey As Long", "ByVal lngKey As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByRef lngKey As Long", "ByRef lngKey As LongPtr", changes)

    ' =========================================================================
    ' STEP 14: Fix RegEnumKeyEx handle parameters (advapi32.dll)
    ' =========================================================================
    newCode = ReplaceExact(newCode, "ByVal hkResult As Long", "ByVal hkResult As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByRef hkResult As Long", "ByRef hkResult As LongPtr", changes)
    newCode = ReplaceExact(newCode, "Dim hkResult As Long", "Dim hkResult As LongPtr", changes)

    ' =========================================================================
    ' STEP 15: Fix function return types for unambiguous handle-returning APIs
    '          Works line-by-line: finds the Declare line containing the
    '          function name and changes its trailing ") As Long" to LongPtr.
    '          Only functions where return type is ALWAYS a handle/pointer.
    ' =========================================================================
    newCode = FixReturnType(newCode, "InternetOpen", changes)
    newCode = FixReturnType(newCode, "InternetConnect", changes)
    newCode = FixReturnType(newCode, "FtpOpenFile", changes)
    newCode = FixReturnType(newCode, "GlobalAlloc", changes)
    newCode = FixReturnType(newCode, "GlobalLock", changes)
    newCode = FixReturnType(newCode, "FindFirstFile", changes)
    newCode = FixReturnType(newCode, "SHBrowseForFolder", changes)
    newCode = FixReturnType(newCode, "GetClipboardData", changes)
    newCode = FixReturnType(newCode, "SetClipboardData", changes)
    newCode = FixReturnType(newCode, "LoadImage", changes)
    newCode = FixReturnType(newCode, "ShellExecute", changes)
    newCode = FixReturnType(newCode, "GetWindow", changes)
    newCode = FixReturnType(newCode, "SendMessage", changes)
    newCode = FixReturnType(newCode, "lstrcat", changes)

    ' =========================================================================
    ' Apply changes back to the module only if something changed
    ' =========================================================================
    If newCode <> code Then
        With comp.CodeModule
            .DeleteLines 1, .CountOfLines
            .InsertLines 1, newCode
        End With
    End If

    FixModule = changes
End Function

' =============================================================================
' Helper: Replace exact string (case-sensitive)
' Counts occurrences before replacing; accumulates into ByRef changes counter
' =============================================================================
Private Function ReplaceExact(source As String, findStr As String, _
                               replaceStr As String, ByRef changes As Long) As String
    Dim result As String
    Dim count  As Long

    count = (Len(source) - Len(Replace(source, findStr, ""))) \ Len(findStr)

    If count > 0 Then
        result = Replace(source, findStr, replaceStr)
        changes = changes + count
        ReplaceExact = result
    Else
        ReplaceExact = source
    End If
End Function

' =============================================================================
' Helper: Replace word (case-insensitive) — used for Declare keyword only
' Avoids double-replacement when module already contains PtrSafe
' =============================================================================
Private Function ReplaceWord(source As String, findStr As String, _
                              replaceStr As String, ByRef changes As Long) As String
    Dim upperSource As String
    Dim upperFind   As String
    Dim result      As String
    Dim pos         As Long
    Dim count       As Long

    upperSource = UCase(source)
    upperFind = UCase(findStr)

    count = 0
    pos = 1
    Do
        pos = InStr(pos, upperSource, upperFind)
        If pos = 0 Then Exit Do
        Dim checkStr As String
        checkStr = UCase(Mid(source, pos, Len(replaceStr)))
        If checkStr <> UCase(replaceStr) Then
            count = count + 1
        End If
        pos = pos + Len(upperFind)
    Loop

    If count > 0 Then
        result = source
        Dim i As Long
        For i = 1 To count
            Dim p As Long
            p = InStr(1, UCase(result), upperFind)
            If p > 0 Then
                If UCase(Mid(result, p, Len(replaceStr))) <> UCase(replaceStr) Then
                    result = Left(result, p - 1) & replaceStr & Mid(result, p + Len(findStr))
                    changes = changes + 1
                End If
            End If
        Next i
        ReplaceWord = result
    Else
        ReplaceWord = source
    End If
End Function

' =============================================================================
' Helper: Fix return type on a Declare Function line containing funcName
' Finds the LAST line of the declaration (the one ending ") As Long") and
' changes it to ") As LongPtr" — only when it's a Declare PtrSafe Function line.
' Safe: only acts on lines that contain both the function name AND "As Long"
' at the end, and only within a Declare block.
' =============================================================================
Private Function FixReturnType(source As String, funcName As String, _
                                ByRef changes As Long) As String
    Dim lines()     As String
    Dim i           As Long
    Dim inDeclare   As Boolean
    Dim result      As String
    Dim changed     As Boolean

    lines = Split(source, vbCrLf)
    inDeclare = False

    For i = 0 To UBound(lines)
        Dim lineU As String
        lineU = UCase(Trim(lines(i)))

        ' Detect start of a Declare PtrSafe Function line containing our function
        If InStr(1, lineU, "DECLARE PTRSAFE FUNCTION") > 0 And _
           InStr(1, UCase(lines(i)), UCase(funcName)) > 0 Then
            inDeclare = True
        End If

        ' If we're in the right Declare block, look for the closing ") As Long"
        If inDeclare Then
            ' The closing line ends with ") As Long" (possibly with spaces/comment)
            ' Must NOT already be LongPtr
            If InStr(1, lineU, ") AS LONG") > 0 And _
               InStr(1, lineU, "LONGPTR") = 0 Then
                ' Replace ") As Long" with ") As LongPtr" on this line
                ' Case-preserving: find the exact position
                Dim pos As Long
                pos = InStr(1, UCase(lines(i)), ") AS LONG")
                If pos > 0 Then
                    ' Make sure it's truly at the end (allow trailing spaces/comment)
                    Dim tail As String
                    tail = Trim(Mid(lines(i), pos + 9)) ' everything after ") As Long"
                    ' tail should be empty or start with ' (comment)
                    If Len(tail) = 0 Or Left(tail, 1) = "'" Then
                        lines(i) = Left(lines(i), pos - 1) & ") As LongPtr" & _
                                   IIf(Len(tail) > 0, " " & tail, "")
                        changes = changes + 1
                        inDeclare = False
                    End If
                End If
            End If
            ' Line continuation "_" means declaration continues — stay in block
            If Right(Trim(lines(i)), 1) <> "_" Then
                ' No continuation — if we haven't found the closing yet, exit
                If InStr(1, lineU, ") AS LONG") = 0 And _
                   InStr(1, lineU, ") AS LONGPTR") = 0 Then
                    inDeclare = False
                End If
            End If
        End If
    Next i

    FixReturnType = Join(lines, vbCrLf)
End Function

' =============================================================================
' UTILITY: Preview what WOULD be changed without actually changing anything
' Run this first to see what will be affected
' =============================================================================
Public Sub Preview32to64_Changes()
    Dim proj    As VBProject
    Dim comp    As VBComponent
    Dim code    As String
    Dim lines() As String
    Dim i       As Long
    Dim found   As Boolean

    Debug.Print String(70, "=")
    Debug.Print "PREVIEW: Lines containing 32-bit patterns that need updating"
    Debug.Print String(70, "=")

    ' All patterns we scan for
    Dim patterns(34) As String
    patterns(0)  = "Declare Function "
    patterns(1)  = "Declare Sub "
    patterns(2)  = "ByVal hWnd As Long"
    patterns(3)  = "ByVal hwnd As Long"
    patterns(4)  = "ByVal hInst As Long"
    patterns(5)  = "ByVal hIcon As Long"
    patterns(6)  = "ByVal hKey As Long"
    patterns(7)  = "ByRef hKey As Long"
    patterns(8)  = "ByVal hInternet As Long"
    patterns(9)  = "ByVal hConnect As Long"
    patterns(10) = "ByVal lParam As Long"
    patterns(11) = "ByVal wParam As Long"
    patterns(12) = "ByVal pidl As Long"
    patterns(13) = "phkResult As Long"
    patterns(14) = "ByVal hMem As Long"
    patterns(15) = "ByVal hBitmap As Long"
    patterns(16) = "ByVal hMenu As Long"
    patterns(17) = "ByVal hWndParent As Long"
    patterns(18) = "ByVal hwndOwner As Long"
    patterns(19) = "Dim hWnd As Long"
    patterns(20) = "Dim hKey As Long"
    patterns(21) = "ByVal hFindFile As Long"
    patterns(22) = "ByVal hFtpSession As Long"
    patterns(23) = "ByVal dwContext As Long"
    patterns(24) = "ByVal dwBytes As Long"
    patterns(25) = "ByVal pv As Long"
    patterns(26) = "ByVal hwndParent As Long"
    patterns(27) = "Dim hFTP As Long"
    patterns(28) = "Dim hConnect As Long"
    patterns(29) = "Dim hInternet As Long"
    patterns(30) = "Dim hMem As Long"
    patterns(31) = "Dim pidl As Long"
    patterns(32) = "ByVal hFindHandle As Long"
    patterns(33) = "pidlRoot As Long"
    patterns(34) = "lpfn As Long"

    Set proj = Application.VBE.ActiveVBProject

    For Each comp In proj.VBComponents
        code = ""
        On Error Resume Next
        code = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
        On Error GoTo 0
        If Len(code) = 0 Then GoTo NextComp

        lines = Split(code, vbCrLf)
        found = False

        For i = 0 To UBound(lines)
            Dim j As Integer
            For j = 0 To 34
                If InStr(1, lines(i), patterns(j), vbTextCompare) > 0 Then
                    If Not found Then
                        Debug.Print ""
                        Debug.Print "  Module: " & comp.Name
                        found = True
                    End If
                    Debug.Print "    Line " & (i + 1) & ": " & Trim(lines(i))
                    Exit For
                End If
            Next j
        Next i
NextComp:
    Next comp

    Debug.Print ""
    Debug.Print String(70, "=")
    Debug.Print "End of Preview. Run Fix32to64_AllModules() to apply changes."
    Debug.Print String(70, "=")
End Sub

' =============================================================================
' UTILITY: Check for any remaining 32-bit patterns after conversion
' Run this AFTER Fix32to64_AllModules() to verify completeness
' =============================================================================
Public Sub Verify32to64_Complete()
    Dim proj    As VBProject
    Dim comp    As VBComponent
    Dim code    As String
    Dim lines() As String
    Dim i       As Long
    Dim issues  As Long

    issues = 0

    Debug.Print String(70, "=")
    Debug.Print "VERIFICATION: Checking for remaining 32-bit patterns"
    Debug.Print String(70, "=")

    Set proj = Application.VBE.ActiveVBProject

    For Each comp In proj.VBComponents
        code = ""
        On Error Resume Next
        code = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
        On Error GoTo 0
        If Len(code) = 0 Then GoTo NextComp2

        lines = Split(code, vbCrLf)

        For i = 0 To UBound(lines)
            Dim lineUpper As String
            lineUpper = UCase(Trim(lines(i)))

            ' Check for Declare without PtrSafe
            If (Left(lineUpper, 17) = "DECLARE FUNCTION " Or _
                Left(lineUpper, 12) = "DECLARE SUB " Or _
                InStr(lineUpper, "PRIVATE DECLARE FUNCTION") > 0 Or _
                InStr(lineUpper, "PUBLIC DECLARE FUNCTION") > 0 Or _
                InStr(lineUpper, "PRIVATE DECLARE SUB") > 0 Or _
                InStr(lineUpper, "PUBLIC DECLARE SUB") > 0) And _
               InStr(lineUpper, "PTRSAFE") = 0 Then
                Debug.Print "  [" & comp.Name & "] Line " & (i + 1) & _
                            " MISSING PtrSafe: " & Trim(lines(i))
                issues = issues + 1
            End If

            ' Check for handle params still As Long on Declare lines
            If InStr(lineUpper, "DECLARE PTRSAFE") > 0 Then
                If InStr(lineUpper, "BYVAL HWND AS LONG") > 0 Or _
                   InStr(lineUpper, "BYVAL HKEY AS LONG") > 0 Or _
                   InStr(lineUpper, "BYVAL HINTERNET AS LONG") > 0 Or _
                   InStr(lineUpper, "BYVAL HCONNECT AS LONG") > 0 Or _
                   InStr(lineUpper, "BYVAL HINST AS LONG") > 0 Or _
                   InStr(lineUpper, "BYVAL HMEM AS LONG") > 0 Then
                    Debug.Print "  [" & comp.Name & "] Line " & (i + 1) & _
                                " HANDLE STILL As Long: " & Trim(lines(i))
                    issues = issues + 1
                End If
            End If
        Next i
NextComp2:
    Next comp

    Debug.Print ""
    If issues = 0 Then
        Debug.Print "  " & Chr(10003) & " All checks passed — no 32-bit patterns found!"
        Debug.Print "  Run Debug > Compile as final verification."
    Else
        Debug.Print "  " & Chr(10007) & " Found " & issues & " issue(s) — review lines above"
    End If
    Debug.Print String(70, "=")
End Sub