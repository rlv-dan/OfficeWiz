Attribute VB_Name = "modBrowseForFolders"
'**************************************
' Name: Browse For Folder Version 3 Final
'     l
' Description:Browse For Folder, now work
' in all version of windows 98/ME/2000/XP
'
' By: Serge Lachapelle
'
'This code is copyrighted and has' limited warranties.Please see http://w
'     ww.Planet-Source-Code.com/vb/scripts/Sho
'     wCode.asp?txtCodeId=49326&lngWId=1'for details.
'**************************************

'   My Notes (Dan):
'       Added code to translate links & NetHood links to target paths (uses GetShortcutTarget)
'           On v6 (XP) shortcuts are translated automatically.
'           GetShortcutTarget: Requires Windows Scripting, but should work (returns selected lnk file or nethood folder) if not installed
'           GetShortcutTarget: Inlcuded as public function. Example: MsgBox GetShortcutTarget("c:\shortcut.lnk")
'       Added code to hide 'New Folder' button (v6 only, new style only)
'       Removed BIF_STATUSTEXT since I didn't think it looked good. (BIF_STATUSTEXT is old ui style only)
'       Shell (what SHBrowseForFolder uses): v4=Win95, v4.71=Win98 (or Win95 with later Internet Explorer installer), v5=WinMe/2k (perhaps IE6?), v6=XP.
'       New style (resize, new folder button, and more ) is v5 only. (IE6?)
'       OwnerForm (first argument) is optional, but should be used (works like vbModal)
'       Setting Ok button caption only works in XP (v6 only?)

'   Examples:
'       MsgBox BrowseForFolder(Form1, "Title", , , "c:\temp", , , "OkButton")
'       MsgBox BrowseForFolder(, , ROOTDIR_CUSTOM, "c:\", "c:\windows", False, True)

'       Include NetHood path to be sure that shortcuts are translated correctly on pre-XP systems
'       MsgBox BrowseForFolder(Me, , , , , , , , , sysFolders.NetHood)



Option Explicit


Private Type SH_ITEM_ID
    cb As Long
    abID As Byte
End Type


Private Type ITEMIDLIST
    mkid As SH_ITEM_ID
End Type


Private Type BrowseInfo
    hwndOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type


Public Enum ROOTDIR_ID
    ROOTDIR_CUSTOM = -1
    ROOTDIR_ALL = &H0
    ROOTDIR_MY_COMPUTER = &H11
    ROOTDIR_DRIVES = &H11
    ROOTDIR_ALL_NETWORK = &H12
    ROOTDIR_NETWORK_COMPUTERS = &H3D
    ROOTDIR_WORKGROUP = &H3D
    ROOTDIR_USER = &H28
    ROOTDIR_USER_DESKTOP = &H10
    ROOTDIR_USER_MY_DOCUMENTS = &H5
    ROOTDIR_USER_START_MENU = &HB
    ROOTDIR_USER_START_MENU_PROGRAMS = &H2
    ROOTDIR_USER_START_MENU_PROGRAMS_STARTUP = &H7
    ROOTDIR_COMMON_DESKTOP = &H19
    ROOTDIR_COMMON_DOCUMENTS = &H2E
    ROOTDIR_COMMON_START_MENU = &H16
    ROOTDIR_COMMON_START_MENU_PROGRAMS = &H17
    ROOTDIR_COMMON_START_MENU_PROGRAMS_STARTUP = &H18
    ROOTDIR_WINDOWS = &H24
    ROOTDIR_SYSTEM = &H25
    ROOTDIR_FONTS = &H14
    ROOTDIR_PROGRAM_FILES = &H26
    ROOTDIR_PROGRAM_FILES_COMMON_FILES = &H2B
End Enum


Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128 ' Maintenance string For PSS usage
End Type
    
Private Const MAX_PATH = 260
Private Const WM_USER = &H400
Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SELCHANGED = 2
Private Const BFFM_SETSTATUSTEXT = (WM_USER + 100)
Private Const BFFM_SETSELECTION = (WM_USER + 102)
Private Const BFFM_SETOKTEXT = (WM_USER + 105)
Private Const BFFM_ENABLEOK = (WM_USER + 101)

Private Const BIF_DEFAULT = &H0
Private Const BIF_RETURNONLYFSDIRS = &H1    'Only return file system directories. If the user selects folders that are not part of the file system, the OK button is greyed out.
Private Const BIF_DONTGOBELOWDOMAIN = &H2   'Do not include network folders below the domain level.
Private Const BIF_STATUSTEXT = &H4          'Include a status area in the dialog box. 'The callback function can set the status text by sending messages to the dialog box. Not With BIF_NEWDIALOGSTYLE
Private Const BIF_RETURNFSANCESTORS = &H8   'Only return file system ancestors. If the user selects anything other than a file system ancestor, the OK button is greyed out.
Private Const BIF_EDITBOX = &H10            'Include an edit control so the user can type the name of an item.
Private Const BIF_VALIDATE = &H20           'If the user types an invalid name into the edit box, the browse dialog will call the application's BrowseCallbackProc with the BFFM_VALIDATEFAILED message. This flag is ignored if BIF_EDITBOX is not specified. Use With BIF_EDITBOX or BIF_USENEWUI
Private Const BIF_NEWDIALOGSTYLE = &H40     ' Use OleInitialize before
Private Const BIF_USENEWUI = &H50           'Use the new user-interface. Setting this flag provides the user with a larger dialog box that can be resized. It has several new capabilities including: drag and drop capability within the dialog box, reordering, context menus, new folders, delete, and other context menu commands. To use this flag, you must call OleInitialize or CoInitialize before calling SHBrowseForFolder. WinMe/2k+ only(?) ' = (BIF_NEWDIALOGSTYLE + BIF_EDITBOX)
Private Const BIF_BROWSEINCLUDEURLS = &H80
Private Const BIF_UAHINT = &H100            'Use With BIF_NEWDIALOGSTYLE, add Usage Hint if no EditBox
Private Const BIF_NONEWFOLDERBUTTON = &H200
Private Const BIF_NOTRANSLATETARGETS = &H400
Private Const BIF_BROWSEFORCOMPUTER = &H1000    'Only return computers. If the user selects anything other than a computer, the OK button is greyed out.
Private Const BIF_BROWSEFORPRINTER = &H2000     'Only return printers.'If the user selects anything other than a printer, the OK button is greyed out.
Private Const BIF_BROWSEINCLUDEFILES = &H4000   'Display files as well as folders.
Private Const BIF_SHAREABLE = &H8000        ' use With BIF_NEWDIALOGSTYLE
' IShellFolder's ParseDisplayName member


'More info from MSDN:
'BIF_BROWSEFORCOMPUTER      'Only return computers. If the user selects anything other than a computer, the OK button is grayed.
'BIF_BROWSEFORPRINTER       'Only allow the selection of printers. If the user selects anything other than a printer, the OK button is grayed.  'In Microsoft Windows XP, the best practice is to use an XP-style dialog, setting the root of the dialog to the Printers and Faxes folder (CSIDL_PRINTERS).
'BIF_BROWSEINCLUDEFILES     'Version 4.71. The browse dialog box will display files as well as folders.
'BIF_BROWSEINCLUDEURLS      'Version 5.0. The browse dialog box can display URLs. The BIF_USENEWUI and BIF_BROWSEINCLUDEFILES flags must also be set. If these three flags are not set, the browser dialog box will reject URLs. Even when these flags are set, the browse dialog box will only display URLs if the folder that contains the selected item supports them. When the folder's IShellFolder::GetAttributesOf method is called to request the selected item's attributes, the folder must set the SFGAO_FOLDER attribute flag. Otherwise, the browse dialog box will not display the URL.
'BIF_DONTGOBELOWDOMAIN      'Do not include network folders below the domain level in the dialog box's tree view control.
'BIF_EDITBOX                'Version 4.71. Include an edit control in the browse dialog box that allows the user to type the name of an item.
'BIF_NEWDIALOGSTYLE         'Version 5.0. Use the new user interface. Setting this flag provides the user with a larger dialog box that can be resized. The dialog box has several new capabilities including: drag-and-drop capability within the dialog box, reordering, shortcut menus, new folders, delete, and other shortcut menu commands. To use this flag, you must call OleInitialize or CoInitialize before calling SHBrowseForFolder.
'BIF_NONEWFOLDERBUTTON      'Version 6.0. Do not include the New Folder button in the browse dialog box.
'BIF_NOTRANSLATETARGETS     'Version 6.0. When the selected item is a shortcut, return the PIDL of the shortcut itself rather than its target.
'BIF_RETURNFSANCESTORS      'Only return file system ancestors. An ancestor is a subfolder that is beneath the root folder in the namespace hierarchy. If the user selects an ancestor of the root folder that is not part of the file system, the OK button is grayed.
'BIF_RETURNONLYFSDIRS       'Only return file system directories. If the user selects folders that are not part of the file system, the OK button is grayed.
'BIF_SHAREABLE              'Version 5.0. The browse dialog box can display shareable resources on remote systems. It is intended for applications that want to expose remote shares on a local system. The BIF_NEWDIALOGSTYLE flag must also be set.
'BIF_STATUSTEXT             'Include a status area in the dialog box. The callback function can set the status text by sending messages to the dialog box. This flag is not supported when BIF_NEWDIALOGSTYLE is specified.
'BIF_UAHINT                 'Version 6.0. When combined with BIF_NEWDIALOGSTYLE, adds a usage hint to the dialog box in place of the edit box. BIF_EDITBOX overrides this flag.
'BIF_USENEWUI               'Version 5.0. Use the new user interface, including an edit box. This flag is equivalent to BIF_EDITBOX | BIF_NEWDIALOGSTYLE. To use BIF_USENEWUI, you must call OleInitialize or CoInitialize before calling SHBrowseForFolder.
'BIF_VALIDATE               'Version 4.71. If the user types an invalid name into the edit box, the browse dialog box will call the application's BrowseCallbackProc with the BFFM_VALIDATEFAILED message. This flag is ignored if BIF_EDITBOX is not specified.


'     function should be used instead.


Private Declare Function SHSimpleIDListFromPath Lib "shell32.dll" Alias "#162" (ByVal szPath As String) As Long
    'Private Declare Function SHILCreateFrom
    '     Path Lib "shell32.dll" (ByVal pszPath As
    '     Long, ByRef ppidl As Long, ByRef rgflnOu
    '     t As Long) As Long


Private Declare Function SHGetPathFromIDList Lib "shell32.dll" (ByVal pidList As Long, ByVal lpBuffer As String) As Long


Private Declare Function SHBrowseForFolder Lib "shell32.dll" (lpbi As BrowseInfo) As Long


Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)


Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long


Private Declare Sub OleInitialize Lib "ole32.dll" (pvReserved As Any)


Private Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long


Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long


Private Declare Function SendMessage2 Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


Private Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
    Private m_CurrentDirectory As String
    Private OK_BUTTON_TEXT As String
    '


Private Function isNT2000XP() As Boolean
    Dim lpv As OSVERSIONINFO
    lpv.dwOSVersionInfoSize = Len(lpv)
    GetVersionEx lpv


    If lpv.dwPlatformId = 2 Then
        isNT2000XP = True
    Else
        isNT2000XP = False
    End If
End Function


Private Function isME2KXP() As Boolean
    Dim lpv As OSVERSIONINFO
    lpv.dwOSVersionInfoSize = Len(lpv)
    GetVersionEx lpv
    If ((lpv.dwPlatformId = 2) And (lpv.dwMajorVersion >= 5)) Or _
    ((lpv.dwPlatformId = 1) And (lpv.dwMajorVersion >= 4) And (lpv.dwMinorVersion >= 90)) Then
    isME2KXP = True
Else
    isME2KXP = False
End If
End Function


Private Function GetPIDLFromPath(spath As String) As Long
    ' Return the pidl to the path supplied b
    '     y calling the undocumented API #162


    If isNT2000XP Then
        GetPIDLFromPath = SHSimpleIDListFromPath(StrConv(spath, vbUnicode))
    Else
        GetPIDLFromPath = SHSimpleIDListFromPath(spath)
    End If
End Function


Private Function GetSpecialFolderID(ByVal csidl As ROOTDIR_ID) As Long
    Dim IDL As ITEMIDLIST, r As Long
    r = SHGetSpecialFolderLocation(ByVal 0&, csidl, IDL)


    If r = 0 Then
        GetSpecialFolderID = IDL.mkid.cb
    Else
        GetSpecialFolderID = 0
    End If
End Function


Private Function GetAddressOfFunction(zAdd As Long) As Long
    GetAddressOfFunction = zAdd
End Function


Private Function BrowseCallbackProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal pData As Long) As Long
    On Local Error Resume Next
    Dim sBuffer As String


    Select Case uMsg
        Case BFFM_INITIALIZED
        SendMessage hwnd, BFFM_SETSELECTION, 1, m_CurrentDirectory
        If OK_BUTTON_TEXT <> vbNullString Then SendMessage2 hwnd, BFFM_SETOKTEXT, 1, StrPtr(OK_BUTTON_TEXT)
        Case BFFM_SELCHANGED
        sBuffer = Space$(MAX_PATH)
        SHGetPathFromIDList lp, sBuffer
        sBuffer = Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)


        If Len(sBuffer) = 0 Then
            SendMessage2 hwnd, BFFM_ENABLEOK, 1, 0
            SendMessage hwnd, BFFM_SETSTATUSTEXT, 1, ""
        Else
            SendMessage hwnd, BFFM_SETSTATUSTEXT, 1, sBuffer
        End If
    End Select
BrowseCallbackProc = 0
End Function


Public Function BrowseForFolder(Optional OwnerForm As Form = Nothing, Optional ByVal Title As String = "", Optional ByVal RootDir As ROOTDIR_ID = ROOTDIR_ALL, Optional ByVal CustomRootDir As String = "", Optional ByVal StartDir As String = "", Optional ByVal NewStyle As Boolean = True, Optional ByVal IncludeFiles As Boolean = False, Optional ByVal OkButtonText As String = "", Optional ByVal HideNewFolderButton As Boolean = False, Optional ByVal NetHoodPath As String = "") As String

    Dim lpIDList As Long, sBuffer As String, tBrowseInfo As BrowseInfo, clRoot As Boolean

    If Len(OkButtonText) > 0 Then
        OK_BUTTON_TEXT = OkButtonText
    Else
        OK_BUTTON_TEXT = vbNullString
    End If
    clRoot = False


    If RootDir = ROOTDIR_CUSTOM Then


        If Len(CustomRootDir) > 0 Then


            If (PathIsDirectory(CustomRootDir) And (Left$(CustomRootDir, 2) <> "\\")) Or (Left$(CustomRootDir, 2) = "\\") Then
                tBrowseInfo.pidlRoot = GetPIDLFromPath(CustomRootDir)
                'SHILCreateFromPath StrPtr(CustomRootDir
                '     ), tBrowseInfo.pidlRoot, ByVal 0&
                clRoot = True
            Else
                tBrowseInfo.pidlRoot = GetSpecialFolderID(ROOTDIR_MY_COMPUTER)
            End If
        Else
            tBrowseInfo.pidlRoot = GetSpecialFolderID(ROOTDIR_ALL)
        End If
    Else
        tBrowseInfo.pidlRoot = GetSpecialFolderID(RootDir)
    End If


    If (Len(StartDir) > 0) Then
        m_CurrentDirectory = StartDir & vbNullChar
    Else
        m_CurrentDirectory = vbNullChar
    End If


    If Len(Title) > 0 Then
        tBrowseInfo.lpszTitle = Title
    Else
        tBrowseInfo.lpszTitle = "Select A Directory"
    End If
    tBrowseInfo.lpfnCallback = GetAddressOfFunction(AddressOf BrowseCallbackProc)
    tBrowseInfo.ulFlags = BIF_RETURNONLYFSDIRS
    If IncludeFiles Then tBrowseInfo.ulFlags = tBrowseInfo.ulFlags + BIF_BROWSEINCLUDEFILES
    
    If NewStyle And isME2KXP Then
        tBrowseInfo.ulFlags = tBrowseInfo.ulFlags + BIF_NEWDIALOGSTYLE + BIF_UAHINT
        If HideNewFolderButton = True Then tBrowseInfo.ulFlags = tBrowseInfo.ulFlags + BIF_NONEWFOLDERBUTTON
        OleInitialize Null ' Initialize OLE and COM
    Else
        'tBrowseInfo.ulFlags = tBrowseInfo.ulFlags + BIF_STATUSTEXT
    End If
    If Not (OwnerForm Is Nothing) Then tBrowseInfo.hwndOwner = OwnerForm.hwnd
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    If clRoot = True Then CoTaskMemFree tBrowseInfo.pidlRoot



    If (lpIDList) Then
        sBuffer = Space$(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        CoTaskMemFree lpIDList
        sBuffer = Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        BrowseForFolder = sBuffer
    
        ' Pre XP, shortcuts are not automatically translated into its target. This code
        ' will check for shortcuts and get their target. There may be other shortcuts
        ' such as NetHood that are not handled by this (PrintHood?), but it seems to
        ' work for all files I have tested so far.
        If NetHoodPath <> "" And Left$(sBuffer, Len(NetHoodPath)) = NetHoodPath Then
            'NetHood shortcut
            BrowseForFolder = GetShortcutTarget(sBuffer & "\target.lnk")
            If BrowseForFolder = "" Then BrowseForFolder = sBuffer
        ElseIf Right$(sBuffer, 4) = ".lnk" Then
            'Normal shortcut (can happen when browsing includes files)
            BrowseForFolder = GetShortcutTarget(sBuffer)
            If BrowseForFolder = "" Then BrowseForFolder = sBuffer
        End If

    Else
        BrowseForFolder = ""
    End If
End Function



'**************************************
' Name: Get Shortcut's Target
' Description:After looking all over on
'     PSC i was unable to find Short, Simple a
'     nd Clean code to get the target path of
'     a window's shortcut (.lnk) file, so here
'     is an easier way.
' By: Michael L. Canejo
'
'This code is copyrighted and has' limited warranties.Please see http://w
'     ww.Planet-Source-Code.com/vb/scripts/Sho
'     wCode.asp?txtCodeId=44596&lngWId=1'for details.'**************************************


Public Function GetShortcutTarget(strPath As String) As String
    'Gets target path from a shortcut file
    On Error GoTo Error_Loading
    Dim wshShell As Object
    Dim wshLink As Object
    Set wshShell = CreateObject("WScript.Shell")
    Set wshLink = wshShell.CreateShortcut(strPath)
    GetShortcutTarget = wshLink.TargetPath
    Set wshLink = Nothing
    Set wshShell = Nothing
    Exit Function
Error_Loading:
    'GetShortcutTarget = "Error occured."
    GetShortcutTarget = ""
    'GetShortcutTarget = strPath
End Function







