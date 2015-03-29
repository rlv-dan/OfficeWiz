Attribute VB_Name = "modCaptureConsole"
Option Explicit

Private Declare Function ExecuteW Lib "CaptureConsole" (ByVal s_CommandLine As Long, ByVal s32_FirstConvert As Long, ByVal s_CurrentDir As Long, ByVal s_Environment As Long, ByVal b_SeparatePipes As Boolean, ByVal s32_Timeout As Long, ByRef s_ApiError As Long, ByRef s_StdOut As Long, ByRef s_StdErr As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal Str As Long) As Long
Private Declare Function lstrcpyW Lib "kernel32" (ByVal Dest As Long, ByVal Src As Long) As Long
Private Declare Function SysFreeString Lib "Oleaut32" (ByVal Bstr As Long) As Long

Public sExecuteOutput As String
Public sExecuteError As String

' Converts a BSTR into a VisualBasic string and frees the allocated memory of the BSTR
Private Function ConvertBSTR(ByVal bs_String As Long) As String

    Dim sString As String
    sString = String(lstrlenW(bs_String), 0)

    lstrcpyW StrPtr(sString), bs_String

    SysFreeString bs_String

    ConvertBSTR = sString
    
End Function

Public Function ExecuteCmd(sCmd As String) As Long


'ExecuteA/W parameters:
's_CommandLine = The entire commandline to be executed. e.g. "C:\Test\Test.bat Param1 Param2"
'u32_FirstConvert = 0 -> Commandline parameter codepage conversion is turned off
'u32_FirstConvert > 0 -> The first commandline parameter to be converted to the DOS codepage (see next chapter for more details)
's_CurrentDir = The current working directory for the Console Application or null if not used.
's_Environment = Additional Environment Variables to be passed to the Console Application: "UserVar1=Value1\nUserVar2=Value2\n"
'You can also override the system variables with your own values. Pass null if not used.
'b_SeparatePipes = true -> capture stdout and stderr with two separate pipes and return them in s_StdOut and s_StdErr
'b_SeparatePipes = false -> capture stdout and stderr with one common pipe and return them in s_StdOut
'u32_Timeout = 0 -> no timeout
'u32_Timeout > 0 -> timeout in milliseconds after which the Console process will be killed


    Dim s_Environ As String
    's_Environ = "UserVariable1=This is UserVariable1" & vbLf & "UserVariable2=This is UserVariable2" & vbLf
    s_Environ = ""
    
    Dim u32_Timeout As Long
    u32_Timeout = 60000 * 2 '2 minuter

    Dim bs_ApiError, bs_StdOut, bs_StdErr, s32_ExitCode As Long
    s32_ExitCode = ExecuteW(StrPtr(sCmd), 0, StrPtr(App.Path), _
                            StrPtr(s_Environ), True, u32_Timeout, bs_ApiError, bs_StdOut, bs_StdErr)

    Dim s_ApiError, s_StdOut, s_StdErr As String
    s_ApiError = ConvertBSTR(bs_ApiError)
    s_StdOut = ConvertBSTR(bs_StdOut)
    s_StdErr = ConvertBSTR(bs_StdErr)
    
    If Len(s_ApiError) > 0 Then
        'MsgBox s_ApiError, vbOKOnly, "API Error"
        sExecuteOutput = ""
        sExecuteError = s_ApiError
    Else
        'MsgBox s_StdOut, vbOKOnly, "StdOut"
        'MsgBox s_StdErr, vbOKOnly, "StdErr"
        sExecuteOutput = s_StdOut
        sExecuteError = s_StdErr
    End If
    
    ExecuteCmd = s32_ExitCode

End Function

