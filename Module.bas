Attribute VB_Name = "Module"
Private Declare Function GetShortPathName Lib "kernel32.dll" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public myFso

Public Type typeDocumentOptions
    
    txtInput As String
    txtOutput As String
    txtBasename As String
    iCreateFolder As Integer    'batch only
    
    sExecuteInfo As String
    iImagesExtracted As Long
    bError As Boolean
    
End Type
Public documentOptions() As typeDocumentOptions
Public iCurrentDocument As Integer

Public isRelease As Boolean
Public bFontSizeWarningShown As Boolean
Public bFailedOnPasswordFile As Boolean

'api open dialog box with multiselect
Public Const OFN_ALLOWMULTISELECT = &H200
Public Const OFN_CREATEPROMPT = &H2000
Public Const OFN_ENABLEHOOK = &H20
Public Const OFN_ENABLETEMPLATE = &H40
Public Const OFN_ENABLETEMPLATEHANDLE = &H80
Public Const OFN_EXPLORER = &H80000                         '  new look commdlg
Public Const OFN_EXTENSIONDIFFERENT = &H400
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_LONGNAMES = &H200000                       '  force long names for 3.x modules
Public Const OFN_NOCHANGEDIR = &H8
Public Const OFN_NODEREFERENCELINKS = &H100000
Public Const OFN_NOLONGNAMES = &H40000                      '  force no long names for 4.x modules
Public Const OFN_NONETWORKBUTTON = &H20000
Public Const OFN_NOREADONLYRETURN = &H8000
Public Const OFN_NOTESTFILECREATE = &H10000
Public Const OFN_NOVALIDATE = &H100
Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_READONLY = &H1
Public Const OFN_SHAREAWARE = &H4000
Public Const OFN_SHAREFALLTHROUGH = 2
Public Const OFN_SHARENOWARN = 1
Public Const OFN_SHAREWARN = 0
Public Const OFN_SHOWHELP = &H10

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
        Flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type
Public Declare Function GetOpenFileName Lib "COMDLG32.DLL" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
'''''''''''''''

' -- ShellAndWait --
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Const STILL_ACTIVE = &H103
Private Const PROCESS_QUERY_INFORMATION = &H400
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
' ------------------

'mouse x/y
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

' --------

' -- Shell Execute (start file with associated program, open folders) --
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_NORMAL = 1&
' ----------------------------------------------------------------------

'hand-muspekare till länkar
Public Const IDC_HAND = 32649&
Public Const IDC_ARROW = 32512&
Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long


Public Sub Main()

    Call PrepareThemeSupport

    frmMain.Show

End Sub

Public Function ShellAndWait(sShell As String, Optional ByRef lExitCode As Long = 0, Optional ByVal eWindowStyle As VBA.VbAppWinStyle = vbNormalFocus, Optional ByRef sError As String, Optional ByVal lTimeOut As Long = 0) As Boolean
    
    'From VbAccelerator.com

    Dim hProcess As Long
    Dim lR As Long
    Dim lTimeStart As Long
    Dim bSuccess As Boolean
    
    On Error GoTo ShellAndWaitError
    
    bSuccess = False

    ' This is v2 which is somewhat more reliable:
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(sShell, eWindowStyle))
    If (hProcess = 0) Then
        sError = "This program could not determine whether the process started.  Please watch the program and check it completes."
        ' Only fail if there is an error - this can happen when the program completes too quickly.
    Else
        bSuccess = True
        lTimeStart = timeGetTime()
        Do
            ' Get the status of the process
            GetExitCodeProcess hProcess, lR
            ' Sleep during wait to ensure the other process gets
            ' processor slice:
            DoEvents: Sleep 100
            If lTimeOut > 0 Then
                If (timeGetTime() - lTimeStart > lTimeOut) Then
                    ' Too long!
                    sError = "The process has timed out."
                    lR = 0
                    bSuccess = False
                End If
            End If
        Loop While lR = STILL_ACTIVE
        lExitCode = lR
    End If

    ShellAndWait = bSuccess
        
    Exit Function


ShellAndWaitError:
    sError = Err.Description
    Exit Function

End Function

Public Function Check_If_Release() As Boolean
    isRelease = False
    Check_If_Release = True
End Function



' Get mouse X coordinates in pixels
'
' If a window handle is passed, the result is relative to the client area
' of that window, otherwise the result is relative to the screen
Function MouseX(Optional ByVal hWnd As Long) As Long
    Dim lpPoint As POINTAPI
    GetCursorPos lpPoint
    If hWnd Then ScreenToClient hWnd, lpPoint
    MouseX = lpPoint.X
End Function

' Get mouse Y coordinates in pixels
'
' If a window handle is passed, the result is relative to the client area
' of that window, otherwise the result is relative to the screen
Function MouseY(Optional ByVal hWnd As Long) As Long
    Dim lpPoint As POINTAPI
    GetCursorPos lpPoint
    If hWnd Then ScreenToClient hWnd, lpPoint
    MouseY = lpPoint.Y
End Function

Public Function GetPathFromFilename(sFileName As String) As String
    n = InStrRev(sFileName, "\")
    If n > 0 Then
        GetPathFromFilename = Mid(sFileName, 1, n)
    End If
End Function


Public Function RemoveFile(ByVal Path) As String

    If Right(Path, 1) = "\" Then Path = (Left(Path, Len(Path) - 1))
    
    For num = Len(Path) To 1 Step -1
        If Mid(Path, num, 1) = "\" Then Exit For
    Next
    'RemoveFile = Right(Path, Len(Path) - num)
    RemoveFile = Left(Path, num)

End Function


Public Function RemovePath(ByVal Path) As String

    If Right(Path, 1) = "\" Then Path = (Left(Path, Len(Path) - 1))
    
    For num = Len(Path) To 1 Step -1
        If Mid(Path, num, 1) = "\" Then Exit For
    Next
    RemovePath = Right(Path, Len(Path) - num)


End Function

Public Function GetFilename(sFileName) As String
    
    Dim pos As Integer
    
    GetFilename = sFileName

    'strip folder
    pos = InStrRev(GetFilename, "\", , vbTextCompare)
    If pos > 1 Then 'Normal
        GetFilename = Right$(GetFilename, Len(GetFilename) - pos)
    End If

    'strip extension
    pos = InStrRev(GetFilename, ".", , vbTextCompare)
    If pos = 0 Then 'No extension
        'GetFilename = sFilename
    ElseIf pos = 1 Then 'Only extension
        GetFilename = ""
    ElseIf pos > 1 Then 'Normal
        GetFilename = Left$(GetFilename, pos - 1)
    End If


End Function

Public Function GetFilenameAndExt(sFileName) As String
    
    Dim pos As Integer
    
    GetFilenameAndExt = sFileName

    'strip folder
    pos = InStrRev(GetFilenameAndExt, "\", , vbTextCompare)
    If pos > 1 Then 'Normal
        GetFilenameAndExt = Right$(GetFilenameAndExt, Len(GetFilenameAndExt) - pos)
    End If


End Function


'Recursive file search
Public Sub RecursiveGetFolderContent(ByVal sFol As String, ByRef OutputList As clsListEmu, bFiles As Boolean, bFolders As Boolean, Optional bRecurse As Boolean = True)

    Dim fld, tFld, tFil
    Dim Filename As String

    Dim Include_Hidden As Integer
    If processHidden = 1 Then Include_Hidden = vbHidden Else Include_Hidden = vbNormal
    If processSystem = 1 Then Include_Hidden = Include_Hidden + vbSystem
    ' below: Dir() does not care about vbReadOnly. all files are returned no matter what...
    'If processWriteProtected = 1 Then Include_Hidden = Include_Hidden + vbReadOnly

    'DoEvents

    On Error GoTo errHandler
    Set fld = myFso.GetFolder(sFol)
    'Filename = Dir(myFso.BuildPath(fld.Path, "*.*"), vbNormal Or vbHidden Or vbSystem Or vbReadOnly)
    Filename = Dir(myFso.BuildPath(fld.Path, "*.*"), Include_Hidden)
    While Len(Filename) <> 0
       If bFiles = True Then OutputList.AddItem myFso.BuildPath(fld.Path, Filename)
       Filename = Dir()  ' Get next file
       'DoEvents
    
        DragDropAddCounter = DragDropAddCounter + 1
        If DragDropAddCounter Mod 90 = 0 Then
            DoEvents
        End If
    
    Wend
    
    If fld.SubFolders.Count > 0 Then
       For Each tFld In fld.SubFolders
          'DoEvents
          If bFolders = True Then OutputList.AddItem tFld.Path
          
          If bRecurse = True Then
            Call RecursiveGetFolderContent(tFld.Path, OutputList, bFiles, bFolders, bRecurse)
          End If
       Next
    End If
    Exit Sub

errHandler:
    Filename = ""
    Resume Next

End Sub

Public Function ShortName(LongPath As String) As String
'***************************************************************************
' Converts Long FileName to Short FileName
'***************************************************************************
    Dim ShortPath As String
    Dim ret As Long
    Const MAX_PATH = 260

    If LongPath = "" Then
        Exit Function
    End If
    ShortPath = Space$(MAX_PATH)
    ret = GetShortPathName(LongPath, ShortPath, MAX_PATH)
    If ret Then
        ShortName = Left$(ShortPath, ret)
    End If
End Function


 Sub SmartCreateFolder(strFolder)
     Dim oFSO: Set oFSO = CreateObject("Scripting.FileSystemObject")
     If oFSO.FolderExists(strFolder) Then
         Exit Sub
     Else
         SmartCreateFolder (oFSO.GetParentFolderName(strFolder))
     End If
     oFSO.CreateFolder (strFolder)
     Set oFSO = Nothing
 End Sub

