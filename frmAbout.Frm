VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " About"
   ClientHeight    =   3180
   ClientLeft      =   4620
   ClientTop       =   4440
   ClientWidth     =   4695
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3180
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   315
      Left            =   45
      TabIndex        =   6
      Top             =   375
      Width           =   570
   End
   Begin VB.Label lblDonate 
      BackStyle       =   0  'Transparent
      Caption         =   "Donate!"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3040
      MouseIcon       =   "frmAbout.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   2190
      Width           =   615
   End
   Begin VB.Image Image3 
      Height          =   870
      Left            =   2040
      Picture         =   "frmAbout.frx":0316
      Top             =   240
      Width           =   945
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "PDF Image Extraction Wizard"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3120
      TabIndex        =   7
      Top             =   240
      Width           =   1335
   End
   Begin VB.Image clickShape 
      Height          =   300
      Left            =   15
      Top             =   15
      Width           =   255
   End
   Begin VB.Label lblReginfo 
      BackStyle       =   0  'Transparent
      Caption         =   "This program is Freeware."
      Height          =   255
      Left            =   2085
      TabIndex        =   5
      Top             =   1440
      Width           =   2340
   End
   Begin VB.Label Label6 
      Caption         =   $"frmAbout.frx":102E
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   3480
      Width           =   2535
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   120
      X2              =   4560
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   120
      X2              =   4560
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Image imgRLV 
      Height          =   2250
      Left            =   120
      MouseIcon       =   "frmAbout.frx":1065
      Picture         =   "frmAbout.frx":136F
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblWeb 
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.rlvision.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      MouseIcon       =   "frmAbout.frx":DAE9
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "All rights reserved"
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2005 RL Vision. All rights reserved."
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   4455
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   1800
      Picture         =   "frmAbout.frx":DDF3
      Top             =   4080
      Width           =   480
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "version 1.0"
      Height          =   195
      Left            =   2085
      TabIndex        =   0
      Top             =   1080
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   2760
      Picture         =   "frmAbout.frx":E235
      Top             =   4080
      Width           =   480
   End
   Begin VB.Shape Shape1 
      Height          =   2490
      Left            =   1920
      Top             =   120
      Width           =   2625
   End
   Begin VB.Label lblStats 
      BackStyle       =   0  'Transparent
      Caption         =   "Support independent software developers! Donate if you like this program:"
      Height          =   615
      Left            =   2085
      TabIndex        =   8
      Top             =   1800
      Width           =   2295
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub clickShape_Click()
    Unload frmAbout
End Sub


Private Sub cmdExit_Click()
    Unload frmAbout
End Sub

Private Sub Form_Click()
    Unload frmAbout
End Sub

Private Sub Form_Load()
    
    'frmMain.Enabled = False
    
    'frmAbout.Caption = " About " & ProgramName
    'frmAbout.lblVersion = "v" & ProgramVersion
    
    'frmAbout.lblTotRenamed = "You have renamed " & totNumberOfFilesRenamed & " files and " & totNumberOfDirsRenamed & " directories" & "!"

    'lblStats = "Files Opened: " & statFilesOpened & Chr(10) _
    & "Lines Scanned: " & statLinesScanned & Chr(10) _
    & "Duplicates Found: " & statDupsFound & Chr(10) _
    & "Unique Dups Found: " & statUniqueDupsFound & Chr(10) _
    & "Duplicates Deleted: " & statDupsDeleted

    'frmAbout.Top = (Screen.Height / 2) - (frmAbout.Height / 2)
    'frmAbout.Left = (Screen.Width / 2) - (frmAbout.Width / 2)
    
    cmdExit.left = frmAbout.width + 1000
    clickShape.BorderStyle = 0
    clickShape.top = 0
    clickShape.left = 0
    clickShape.width = frmAbout.width
    clickShape.height = frmAbout.height
    clickShape.ZOrder
    lblWeb.ZOrder
    lblDonate.ZOrder
    'lblMail.ZOrder

End Sub


Private Sub Form_Unload(Cancel As Integer)

    'frmMain.Enabled = True

End Sub


Private Sub imgRLV_Click()

'Temp = ShellExecute(Me.hwnd, vbNullString, "http://www.rlvision.com", vbNullString, "c:\", SW_SHOWNORMAL)

End Sub

Private Sub lblMail_Click()

'Temp = ShellExecute(Me.hwnd, vbNullString, "mailto:saeden@swipnet.se", vbNullString, "c:\", SW_SHOWNORMAL)
'Temp = ShellExecute(Me.hwnd, vbNullString, "mailto:dan@saeden.cjb.net?Subject=Flash Renamer", vbNullString, "c:\", SW_SHOWNORMAL)
Temp = ShellExecute(Me.hWnd, vbNullString, "mailto:info@rlvision.com?Subject=" & ProgramName & " " & ProgramVersion, vbNullString, "c:\", SW_SHOWNORMAL)
End Sub


Private Sub lblDonate_Click()

    Call ShellExecute(Me.hWnd, "open", "http://www.rlvision.com/flashren/about.asp", vbNullString, vbNullString, SW_NORMAL)

End Sub

Private Sub lblWeb_Click()

'Temp = ShellExecute(Me.hwnd, vbNullString, "http://saeden.cjb.net", vbNullString, "c:\", SW_SHOWNORMAL)
Temp = ShellExecute(Me.hWnd, vbNullString, "http://www.rlvision.com/script/redirect.asp?app=dupli", vbNullString, "c:\", SW_SHOWNORMAL)

End Sub


