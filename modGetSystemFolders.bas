Attribute VB_Name = "modGetSystemFolders"
'**************************************
' Name: getSystemFolders()
'
' Description: get various system folder
' locations and store in public variable
' sysFolders. Works without shfolder.dll (pre IE4?)
'
' By: Dan
'
'**************************************


Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''

Private Type systemFolders
    System As String
    Windows As String
    ProgramFiles As String
    SysRoot As String
    MyDocuments As String
    MyPictures As String
    NetHood As String
    Temp As String
End Type

Public sysFolders As systemFolders

'''''''''''''''''''''''''''''''''''''''''''''

'non-shfolder stuff

Const MAX_PATH = 206

Private Declare Function GetWindowsDirectory Lib "kernel32" _
    Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, _
    ByVal nSize As Long) As Long
    
Private Declare Function GetSystemDirectory Lib "kernel32" _
    Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, _
    ByVal nSize As Long) As Long


'''''''''''''''''''''''''''''''''''''''''''''

'shfolder stuff

Private Const CSIDL_ADMINTOOLS           As Long = &H30   '{user}\Start Menu _                                                        '\Programs\Administrative Tools
Private Const CSIDL_COMMON_ADMINTOOLS    As Long = &H2F   '(all users)\Start Menu\Programs\Administrative Tools
Private Const CSIDL_APPDATA              As Long = &H1A   '{user}\Application Data
Private Const CSIDL_COMMON_APPDATA       As Long = &H23   '(all users)\Application Data
Private Const CSIDL_COMMON_DOCUMENTS     As Long = &H2E   '(all users)\Documents
Private Const CSIDL_COOKIES              As Long = &H21
Private Const CSIDL_HISTORY              As Long = &H22
Private Const CSIDL_INTERNET_CACHE       As Long = &H20   'Internet Cache folder
Private Const CSIDL_LOCAL_APPDATA        As Long = &H1C   '{user}\Local Settings\Application Data (non roaming)
Private Const CSIDL_MYPICTURES           As Long = &H27   'C:\Program Files\My Pictures
Private Const CSIDL_PERSONAL             As Long = &H5    'My Documents
Private Const CSIDL_PROGRAM_FILES        As Long = &H26   'Program Files folder
Private Const CSIDL_PROGRAM_FILES_COMMON As Long = &H2B   'Program Files\Common
Private Const CSIDL_SYSTEM               As Long = &H25   'system folder
Private Const CSIDL_WINDOWS              As Long = &H24   'Windows directory or SYSROOT()
Private Const CSIDL_FLAG_CREATE = &H8000&                 'combine with CSIDL_ value to force
Private Const MAX_PATH_SHFOLDER = 260
'
' Other Special Folder CSIDLs
' -> Some of these may not work with SHGetFolderPath() -> test!
' -> CSIDL_NETHOOD does work, at least on 2000/XP.
'
Private Const CSIDL_ALTSTARTUP As Long = &H1D             'non localized startup
Private Const CSIDL_BITBUCKET As Long = &HA               '{desktop}\Recycle Bin
Private Const CSIDL_CONTROLS As Long = &H3                'My Computer\Control Panel
Private Const CSIDL_DESKTOP As Long = &H0                 '{namespace root}
Private Const CSIDL_DESKTOPDIRECTORY As Long = &H10       '{user}\Desktop
Private Const CSIDL_FAVORITES As Long = &H6               '{user}\Favourites
Private Const CSIDL_FONTS As Long = &H14                  'windows\fonts
Private Const CSIDL_INTERNET As Long = &H1                'Internet virtual folder
Private Const CSIDL_DRIVES As Long = &H11                 'My Computer
Private Const CSIDL_NETHOOD As Long = &H13                '{user}\nethood
Private Const CSIDL_NETWORK As Long = &H12                'Network Neighbourhood
Private Const CSIDL_PRINTERS As Long = &H4                'My Computer\Printers
Private Const CSIDL_PRINTHOOD As Long = &H1B              '{user}\PrintHood
Private Const CSIDL_PROGRAM_FILESX86 As Long = &H2A       'Program Files folder for x86 apps (Alpha)
Private Const CSIDL_PROGRAMS As Long = &H2                'Start Menu\Programs
Private Const CSIDL_PROGRAM_FILES_COMMONX86 As Long = &H2C 'x86 \Program Files\Common on RISC
Private Const CSIDL_RECENT As Long = &H8                  '{user}\Recent
Private Const CSIDL_SENDTO As Long = &H9                  '{user}\SendTo
Private Const CSIDL_STARTMENU As Long = &HB               '{user}\Start Menu
Private Const CSIDL_STARTUP As Long = &H7                 'Start Menu\Programs\Startup
Private Const CSIDL_SYSTEMX86 As Long = &H29              'system folder for x86 apps (Alpha)
Private Const CSIDL_TEMPLATES As Long = &H15
Private Const CSIDL_PROFILE As Long = &H28                'user's profile folder
Private Const CSIDL_COMMON_ALTSTARTUP As Long = &H1E      'non localized common startup
Private Const CSIDL_COMMON_DESKTOPDIRECTORY As Long = &H19 '(all users)\Desktop
Private Const CSIDL_COMMON_FAVORITES As Long = &H1F       '(all users)\Favourites
Private Const CSIDL_COMMON_PROGRAMS As Long = &H17        '(all users)\Programs
Private Const CSIDL_COMMON_STARTMENU As Long = &H16       '(all users)\Start Menu
Private Const CSIDL_COMMON_STARTUP As Long = &H18         '(all users)\Startup
Private Const CSIDL_COMMON_TEMPLATES As Long = &H2D       '(all users)\Templates
'                                                          'create on SHGetSpecialFolderLocation()
'Private Const CSIDL_FLAG_DONT_VERIFY = &H4000             'combine with CSIDL_ value to force
'                                                          'create on SHGetSpecialFolderLocation()

Private Const CSIDL_FLAG_MASK = &HFF00                    'mask for all possible flag values
Private Const SHGFP_TYPE_CURRENT = &H0                    'current value for user, verify it exists
Private Const SHGFP_TYPE_DEFAULT = &H1
Private Const S_OK = 0
Private Const S_FALSE = 1
Private Const E_INVALIDARG = &H80070057                   ' Invalid CSIDL Value

Private Declare Function SHGetFolderPath Lib "shfolder" _
        Alias "SHGetFolderPathA" (ByVal hwndOwner As Long, _
        ByVal nFolder As Long, ByVal hToken As Long, _
        ByVal dwFlags As Long, ByVal pszPath As String) As Long



Public Sub getSystemFolders(mySysFolders As systemFolders)

    On Error GoTo errHandler

    Dim errNum As Integer
    errNum = 1

    Dim lHandle As Long
    Dim lResult As Long
    Dim sBuffer As String
    '
    ' Get the Windows folder.
    '
    sBuffer = Space$(MAX_PATH)
    mySysFolders.Windows = GetWindowsDirectory(sBuffer, MAX_PATH)
    mySysFolders.Windows = Left$(sBuffer, Len(sBuffer) - 1)
    mySysFolders.Windows = Trim(mySysFolders.Windows)
    mySysFolders.Windows = Replace(mySysFolders.Windows, Chr(0), "")
    If Right(mySysFolders.Windows, 1) <> "\" Then mySysFolders.Windows = mySysFolders.Windows & "\"
    '
    ' Get the System folder.
    '
    sBuffer = Space$(MAX_PATH)
    mySysFolders.System = GetSystemDirectory(sBuffer, MAX_PATH)
    mySysFolders.System = Left$(sBuffer, Len(sBuffer) - 1)
    mySysFolders.System = Trim(mySysFolders.System)
    mySysFolders.System = Replace(mySysFolders.System, Chr(0), "")
    If Right(mySysFolders.System, 1) <> "\" Then mySysFolders.System = mySysFolders.System & "\"

    'assume that sysroot is at same drive as windows folder, usually c:
    mySysFolders.SysRoot = Left(mySysFolders.Windows, 3)

    'get temp folder
    mySysFolders.Temp = Environ("TEMP")

    'alternative way to get temp folder using api. since above seem to work I don't see any point of using the api version:
        'Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
        'Public Const MAX_BUFFER_LENGTH = 256
        '
        'Public Function getTempPathName() As String
        '    Dim strBufferString As String
        '    Dim lngResult As Long
        '    strBufferString = String(MAX_BUFFER_LENGTH, "X")
        '    lngResult = GetTempPath(MAX_BUFFER_LENGTH, strBufferString)
        '    getTempPathName = Mid(strBufferString, 1, lngResult)
        'End Function


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Test is shfolder.dll is present in system to get more info
    'I think shfolder.dll is installed with Internet Explorer 4+

    errNum = 2
    
    Dim ret As String
    ret = Dir(mySysFolders.System & "shfolder.dll", vbHidden + vbReadOnly + vbSystem)
    If ret <> "" Then   'shfolder excists, so use it!
    
        mySysFolders.ProgramFiles = shGetFolder(CSIDL_PROGRAM_FILES)
        If mySysFolders.ProgramFiles = "" Then mySysFolders.ProgramFiles = "c:\Program Files"
        If Right(mySysFolders.ProgramFiles, 1) <> "\" Then mySysFolders.ProgramFiles = mySysFolders.ProgramFiles & "\"
        
        mySysFolders.MyDocuments = shGetFolder(CSIDL_PERSONAL)
        If Right(mySysFolders.MyDocuments, 1) <> "\" Then mySysFolders.MyDocuments = mySysFolders.MyDocuments & "\"

        mySysFolders.MyPictures = shGetFolder(CSIDL_MYPICTURES)
        If Right(mySysFolders.MyPictures, 1) <> "\" Then mySysFolders.MyPictures = mySysFolders.MyPictures & "\"
        
        mySysFolders.NetHood = shGetFolder(CSIDL_NETHOOD)
        If Right(mySysFolders.NetHood, 1) <> "\" Then mySysFolders.NetHood = mySysFolders.NetHood & "\"

    Else
        mySysFolders.ProgramFiles = "c:\Program Files\" 'set default
    End If
        
        
    'final error checking
    If mySysFolders.Windows = "" Then mySysFolders.Windows = "c:\Windows\"
    If mySysFolders.System = "" Then mySysFolders.System = "c:\Windows\System32\"
    If mySysFolders.ProgramFiles = "" Then mySysFolders.ProgramFiles = "c:\Program Files\"
    If mySysFolders.SysRoot = "" Then mySysFolders.SysRoot = "c:\"

        
    Exit Sub

errHandler:
    'MsgBox ("Error " & Err.Number & " (" & Err.Description & ") " & "getSystemFolders (errNum:" & errNum & ")")
    mySysFolders.Windows = "c:\Windows\"
    mySysFolders.System = "c:\Windows\System32\"
    mySysFolders.ProgramFiles = "c:\Program Files\"
    mySysFolders.SysRoot = "c:\"

End Sub

Public Function shGetFolder(myCSIDL) As String

    On Error GoTo errHandler

        Dim strBuffer  As String
        Dim strPath    As String
        Dim lngReturn  As Long
        Dim lngCSIDL   As Long

        lngCSIDL = myCSIDL
        strPath = String(MAX_PATH_SHFOLDER, 0)
        
        '
        ' Get the folder's path. If the
        ' "Create" flag is used, the folder will be created
        ' if it does not exist.
        '
        lngReturn = SHGetFolderPath(0, lngCSIDL, 0, SHGFP_TYPE_CURRENT, strPath)
        'lngReturn = SHGetFolderPath(0, lngCSIDL Or CSIDL_FLAG_CREATE, 0, SHGFP_TYPE_CURRENT, strPath)
        
        Select Case lngReturn
            Case S_OK   'Ok
                shGetFolder = Left$(strPath, InStr(1, strPath, Chr(0)) - 1)

            Case S_FALSE    'Folder Does Not Exist
                shGetFolder = ""
            
            Case E_INVALIDARG   'Folder Not Valid on this OS
                shGetFolder = ""
            
            Case Else   'Folder Not Valid on this OS
                shGetFolder = ""
            
        End Select

    Exit Function
    
errHandler:
   'MsgBox ("Error " & Err.Number & " (" & Err.Description & ") " & "shGetFolder")

End Function



