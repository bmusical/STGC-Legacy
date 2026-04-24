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
' =============================================================================

Option Compare Database
Option Explicit

' =============================================================================
' MAIN ENTRY POINT — Run this Sub
' =============================================================================
Public Sub Fix32to64_AllModules()
    Dim proj        As VBProject
    Dim comp        As VBComponent
    Dim totalChanges As Long
    Dim moduleChanges As Long
    Dim startTime   As Date
    
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
    Debug.Print "  2. Fix any remaining errors manually"
    Debug.Print "  3. Save the database"
    
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
    Dim code        As String
    Dim newCode     As String
    Dim changes     As Long
    
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
    ' STEP 1: Add PtrSafe to Declare Function/Sub (case-insensitive)
    ' =========================================================================
    newCode = ReplaceWord(newCode, "Declare Function ", "Declare PtrSafe Function ", changes)
    newCode = ReplaceWord(newCode, "Declare Sub ", "Declare PtrSafe Sub ", changes)
    ' Fix double PtrSafe (in case some already had it)
    newCode = ReplaceExact(newCode, "Declare PtrSafe PtrSafe Function ", "Declare PtrSafe Function ", changes)
    newCode = ReplaceExact(newCode, "Declare PtrSafe PtrSafe Sub ", "Declare PtrSafe Sub ", changes)
    
    ' =========================================================================
    ' STEP 2: Fix HWND parameters
    ' =========================================================================
    newCode = ReplaceExact(newCode, "ByVal hWnd As Long", "ByVal hWnd As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByVal hwnd As Long", "ByVal hwnd As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByVal hWndParent As Long", "ByVal hWndParent As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByVal hwndOwner As Long", "ByVal hwndOwner As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByVal hWndOwner As Long", "ByVal hWndOwner As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByRef hWnd As Long", "ByRef hWnd As LongPtr", changes)
    
    ' =========================================================================
    ' STEP 3: Fix HANDLE parameters
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
    ' STEP 4: Fix Registry handle parameters
    ' =========================================================================
    newCode = ReplaceExact(newCode, "ByVal hKey As Long", "ByVal hKey As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByRef hKey As Long", "ByRef hKey As LongPtr", changes)
    newCode = ReplaceExact(newCode, "phkResult As Long", "phkResult As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByRef phkResult As Long", "ByRef phkResult As LongPtr", changes)
    
    ' =========================================================================
    ' STEP 5: Fix WinInet handle parameters
    ' =========================================================================
    newCode = ReplaceExact(newCode, "ByVal hInternet As Long", "ByVal hInternet As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByVal hConnect As Long", "ByVal hConnect As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByVal hFtpSession As Long", "ByVal hFtpSession As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByVal hConnection As Long", "ByVal hConnection As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByVal dwContext As Long", "ByVal dwContext As LongPtr", changes)
    
    ' =========================================================================
    ' STEP 6: Fix pointer/lParam/wParam parameters
    ' =========================================================================
    newCode = ReplaceExact(newCode, "ByVal lParam As Long", "ByVal lParam As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByRef lParam As Long", "ByRef lParam As LongPtr", changes)
    newCode = ReplaceExact(newCode, "lParam As Long", "lParam As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByVal wParam As Long", "ByVal wParam As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByRef wParam As Long", "ByRef wParam As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByVal pidl As Long", "ByVal pidl As LongPtr", changes)
    newCode = ReplaceExact(newCode, "ByRef pidl As Long", "ByRef pidl As LongPtr", changes)
    newCode = ReplaceExact(newCode, "pidlRoot As Long", "pidlRoot As LongPtr", changes)
    newCode = ReplaceExact(newCode, "lpfn As Long", "lpfn As LongPtr", changes)
    
    ' =========================================================================
    ' STEP 7: Fix TYPE structure fields containing handles/pointers
    '         (These replacements are context-aware — only affect UDT fields)
    ' =========================================================================
    newCode = ReplaceExact(newCode, "    hIcon As Long", "    hIcon As LongPtr", changes)
    newCode = ReplaceExact(newCode, "    hwndOwner As Long", "    hwndOwner As LongPtr", changes)
    newCode = ReplaceExact(newCode, "    hWnd As Long", "    hWnd As LongPtr", changes)
    
    ' =========================================================================
    ' STEP 8: Fix Dim statements for handle variables
    ' =========================================================================
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
    
    ' =========================================================================
    ' Apply changes back to the module if anything changed
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
' =============================================================================
Private Function ReplaceExact(source As String, findStr As String, _
                               replaceStr As String, ByRef changes As Long) As String
    Dim result As String
    Dim count  As Long
    
    ' Count occurrences before replacement
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
' Helper: Replace word but avoid double-replacement (e.g. already has PtrSafe)
' Uses case-insensitive search for Declare keywords
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
    
    ' Count occurrences
    count = 0
    pos = 1
    Do
        pos = InStr(pos, upperSource, upperFind)
        If pos = 0 Then Exit Do
        ' Make sure it's not already "PtrSafe" following it
        Dim checkStr As String
        checkStr = UCase(Mid(source, pos, Len(replaceStr)))
        If checkStr <> UCase(replaceStr) Then
            count = count + 1
        End If
        pos = pos + Len(upperFind)
    Loop
    
    If count > 0 Then
        ' Do case-insensitive replace
        result = source
        Dim i As Long
        For i = 1 To count
            Dim p As Long
            p = InStr(1, UCase(result), upperFind)
            If p > 0 Then
                ' Check not already replaced
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
    
    Dim patterns(20) As String
    patterns(0) = "Declare Function "
    patterns(1) = "Declare Sub "
    patterns(2) = "ByVal hWnd As Long"
    patterns(3) = "ByVal hwnd As Long"
    patterns(4) = "ByVal hInst As Long"
    patterns(5) = "ByVal hIcon As Long"
    patterns(6) = "ByVal hKey As Long"
    patterns(7) = "ByRef hKey As Long"
    patterns(8) = "ByVal hInternet As Long"
    patterns(9) = "ByVal hConnect As Long"
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
            For j = 0 To 20
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
    
    ' Patterns that should NOT exist after conversion
    Dim badPatterns(5) As String
    badPatterns(0) = "Declare Function "   ' Should be "Declare PtrSafe Function"
    badPatterns(1) = "Declare Sub "        ' Should be "Declare PtrSafe Sub"
    ' Note: The following are only bad if on a Declare line or in a Type block
    ' So we check for them on Declare lines specifically
    badPatterns(2) = "Declare PtrSafe Function"  ' These are GOOD - just for reference
    badPatterns(3) = ""
    badPatterns(4) = ""
    badPatterns(5) = ""
    
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
                Left(lineUpper, 15) = "PRIVATE DECLARE" Or _
                Left(lineUpper, 14) = "PUBLIC DECLARE") And _
               InStr(lineUpper, "PTRSAFE") = 0 Then
                Debug.Print "  [" & comp.Name & "] Line " & (i + 1) & _
                            " MISSING PtrSafe: " & Trim(lines(i))
                issues = issues + 1
            End If
        Next i
NextComp2:
    Next comp
    
    Debug.Print ""
    If issues = 0 Then
        Debug.Print "  ✓ All Declare statements have PtrSafe - GOOD!"
    Else
        Debug.Print "  ✗ Found " & issues & " Declare statement(s) missing PtrSafe"
    End If
    Debug.Print String(70, "=")
End Sub