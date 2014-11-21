Attribute VB_Name = "pubModule"
Public Const STARTF_USESHOWWINDOW = &H1
Public Const EM_SETTABSTOPS = &HCB
Public Const SW_SHOWNORMAL As Long = 1
Public Const INFINITE = &HFFFFFFFF
Public Const WAIT_TIMEOUT As Long = &H102&

Public Enum enSW
    SW_Hide = 0
    SW_NORMAL = 1
    SW_MAXIMIZE = 3
    SW_MINIMIZE = 6
End Enum
Public Type PROCESS_INFORMATION
    hProcess   As Long
    hThread   As Long
    dwProcessId   As Long
    dwThreadId   As Long
End Type
Public Type STARTUPINFO
    cb   As Long
    lpReserved   As String
    lpDesktop   As String
    lpTitle   As String
    dwX   As Long
    dwY   As Long
    dwXSize   As Long
    dwYSize   As Long
    dwXCountChars   As Long
    dwYCountChars   As Long
    dwFillAttribute   As Long
    dwFlags   As Long
    wShowWindow   As Integer
    cbReserved2   As Integer
    lpReserved2   As Byte
    hStdInput   As Long
    hStdOutput   As Long
    hStdError   As Long
End Type
Public Type SECURITY_ATTRIBUTES
    nLength   As Long
    lpSecurityDescriptor   As Long
    bInheritHandle   As Long
End Type
Public Enum enPriority_Class
    NORMAL_PRIORITY_CLASS = &H20
    IDLE_PRIORITY_CLASS = &H40
    HIGH_PRIORITY_CLASS = &H80
End Enum
Public Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Public Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type
Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type
Public Type WIN32_FIND_DATA
  dwFileAttributes  As Long
  ftCreationTime    As FILETIME
  ftLastAccessTime  As FILETIME
  ftLastWriteTime   As FILETIME
  nFileSizeHigh     As Long
  nFileSizeLow      As Long
  dwReserved0       As Long
  dwReserved1       As Long
  cFileName         As String * 260
  cAlternate        As String * 14
End Type
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BROWSEINFO) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Public Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'----------------------------------------------------------------------------------------------

Public Function Delay(ByVal nn As Single)
Dim tm1 As Long, tm2  As Long
tm1 = timeGetTime
Do
    tm2 = timeGetTime
    If (tm2 - tm1) / 1000 > nn Then Exit Do
    DoEvents
Loop
End Function

Public Function sOpenDir(Optional sTitle As String) As String
    Dim bi As BROWSEINFO
    Dim pidl As Long
    Dim sDir As String
    sDir = String(255, vbNullChar)
    With bi
        .hOwner = 0
        .ulFlags = 81
        .pidlRoot = 0
        .lpszTitle = IIf(sTitle <> "", sTitle & vbNullChar, "Select:" & vbNullChar)
    End With
    pidl = SHBrowseForFolder(bi)
    If SHGetPathFromIDList(ByVal pidl, ByVal sDir) Then
        sOpenDir = Left(sDir, InStr(sDir, vbNullChar) - 1)
    Else
        sOpenDir = ""
    End If
End Function

Public Function sOpenFile(Optional sFilter As String, Optional sDir As String, Optional sTitle As String) As String
Dim sfName As String
Dim ofn As OPENFILENAME
ofn.lStructSize = Len(ofn)
ofn.hwndOwner = 0
ofn.hInstance = App.hInstance
ofn.lpstrFilter = IIf(sFilter = "", "*.*" + Chr$(0) + "*.*" + Chr$(0), sFilter)
ofn.lpstrFile = Space$(254)
ofn.nMaxFile = 255
ofn.lpstrFileTitle = Space$(254)
ofn.nMaxFileTitle = 255
ofn.lpstrInitialDir = IIf(sDir = "", CurDir, sDir)
ofn.lpstrTitle = IIf(sTitle = "", "Open", sTitle)
ofn.flags = 0
If (GetOpenFileName(ofn)) Then
    sfName = Trim$(ofn.lpstrFile)
    If Right(sfName, 1) = Chr$(0) Then sfName = Left(sfName, Len(sfName) - 1)
End If
sOpenFile = sfName
End Function

Public Function sSaveFile(Optional sFilter As String, Optional sDir As String, Optional sTitle As String, Optional sName As String) As String
Dim sfName As String
Dim ofn As OPENFILENAME
ofn.lStructSize = Len(ofn)
ofn.hwndOwner = 0
ofn.hInstance = App.hInstance
ofn.lpstrFilter = IIf(sFilter = "", "*.*" + Chr$(0) + "*.*" + Chr$(0), sFilter)
ofn.lpstrFile = IIf(sName = "", Space$(254), sName & Space$(254 - Len(sName)))
ofn.nMaxFile = 255
ofn.lpstrFileTitle = Space$(254)
ofn.nMaxFileTitle = 255
ofn.lpstrInitialDir = IIf(sDir = "", CurDir, sDir)
ofn.lpstrTitle = IIf(sTitle = "", "Open", sTitle)
ofn.flags = 0
If (GetSaveFileName(ofn)) Then
    sfName = Trim$(ofn.lpstrFile)
    If Right(sfName, 1) = Chr$(0) Then sfName = Left(sfName, Len(sfName) - 1)
End If
sSaveFile = sfName
End Function


Public Function ExShell(ByVal App As String, ByVal WorkDir As String, ByVal start_size As enSW, ByVal Priority_Class As enPriority_Class) As Boolean
    Dim pclass As Long
    Dim sinfo As STARTUPINFO
    Dim pinfo As PROCESS_INFORMATION
    Dim sec1 As SECURITY_ATTRIBUTES
    Dim sec2 As SECURITY_ATTRIBUTES
    Dim nResult As Long
    sec1.nLength = Len(sec1)
    sec2.nLength = Len(sec2)
    sinfo.cb = Len(sinfo)
    sinfo.dwFlags = STARTF_USESHOWWINDOW
    sinfo.wShowWindow = start_size
    pclass = Priority_Class
    If CreateProcess(vbNullString, App, sec1, sec2, False, pclass, 0&, WorkDir, sinfo, pinfo) Then
        Do
            DoEvents
            nResult = WaitForSingleObject(pinfo.hProcess, 100)
        Loop Until nResult <> WAIT_TIMEOUT
        CloseHandle (pinfo.hProcess)
    End If
End Function

Public Function ShortName(LongPath As String) As String
Dim ShortPath As String
Const MAX_PATH = 260
Dim ret&
ShortPath = Space$(MAX_PATH)
ret& = GetShortPathName(LongPath, ShortPath, MAX_PATH)
If ret& Then
ShortName = Left$(ShortPath, ret&)
End If
End Function

Function getAbnormal(inArray(), ByRef iSame, ByRef iDiff)
Dim B As Boolean
B = False
For i = 1 To UBound(inArray)
    If inArray(0) = inArray(i) Then
        iSame = inArray(0)
        B = True
        Exit For
    End If
Next
If B = True Then
    For i = 1 To UBound(inArray)
        If inArray(0) <> inArray(i) Then
            iDiff = inArray(i)
            Exit For
        End If
    Next
Else
    iDiff = inArray(0)
    iSame = inArray(1)
End If
End Function

Public Function Digitrim(i As Double) As String
If i >= 1000 Then Digitrim = Format(i, "0.00e+0#")
If i < 1000 And i >= 100 Then Digitrim = Format(i, "0")
If i < 100 And i >= 10 Then Digitrim = Format(i, "0.0")
If i < 10 And i >= 1 Then Digitrim = Format(i, "0.00")
If i < 1 And i >= 0.1 Then Digitrim = Format(i, "0.000")
If i < 0.1 And i >= 0.01 Then Digitrim = Format(i, "0.0000")
If i < 0.01 And i >= 0.001 Then Digitrim = Format(i, "0.00000")
If i < 0.001 Then Digitrim = Format(i, "0.00e-0#")
If i = 0 Then Digitrim = "0"
End Function

