VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Office Image Extraction Wizard"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   OLEDropMode     =   1  'Manual
   Picture         =   "frmMain.frx":0E42
   ScaleHeight     =   5115
   ScaleWidth      =   6855
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame picFrameWelcome 
      BackColor       =   &H00FFFFFF&
      Height          =   4455
      Left            =   2160
      OLEDropMode     =   1  'Manual
      TabIndex        =   14
      Top             =   0
      Width           =   4695
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Developed by RL Vision © 2010-2015"
         Height          =   195
         Left            =   120
         OLEDropMode     =   1  'Manual
         TabIndex        =   64
         Top             =   900
         Width           =   2700
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00808080&
         X1              =   120
         X2              =   4200
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Freeware (Open Source)"
         Height          =   195
         Left            =   120
         OLEDropMode     =   1  'Manual
         TabIndex        =   62
         Top             =   1500
         Width           =   1740
      End
      Begin VB.Label lblLinkButton 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.rlvision.com"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   2
         Left            =   120
         OLEDropMode     =   1  'Manual
         TabIndex        =   15
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label lblProgramInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "This program allows you to easily extract raw images from various documents types and e-book file formats. Click 'next' to begin."
         Height          =   615
         Left            =   120
         OLEDropMode     =   1  'Manual
         TabIndex        =   17
         Top             =   3525
         Width           =   3975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome to Office Image Extraction Wizard!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         OLEDropMode     =   1  'Manual
         TabIndex        =   16
         Top             =   120
         Width           =   4335
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   375
      Left            =   5640
      TabIndex        =   8
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next >"
      Default         =   -1  'True
      Height          =   375
      Left            =   3975
      TabIndex        =   7
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "< Back"
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Frame picFrameTop 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   0
      OLEDropMode     =   1  'Manual
      TabIndex        =   10
      Top             =   0
      Width           =   6855
      Begin VB.PictureBox picSplashTop 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   900
         Left            =   5840
         Picture         =   "frmMain.frx":0ECD
         ScaleHeight     =   900
         ScaleWidth      =   900
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   100
         Width           =   900
      End
      Begin VB.Label lblTopTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Title"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         OLEDropMode     =   1  'Manual
         TabIndex        =   13
         Top             =   240
         Width           =   390
      End
      Begin VB.Label lblTopText 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   615
         Left            =   600
         OLEDropMode     =   1  'Manual
         TabIndex        =   12
         Top             =   480
         Width           =   5055
      End
   End
   Begin VB.PictureBox picSplash 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   0
      Picture         =   "frmMain.frx":2476
      ScaleHeight     =   4335
      ScaleWidth      =   2490
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   2490
   End
   Begin VB.Frame picFrameBegin 
      Height          =   3375
      Left            =   0
      TabIndex        =   36
      Top             =   1080
      Width           =   6855
      Begin VB.CheckBox chkSkipReadyPage 
         Caption         =   "Skip this page in the future."
         Height          =   195
         Left            =   1200
         TabIndex        =   39
         Top             =   2040
         Width           =   2895
      End
      Begin VB.Image Image3 
         Height          =   405
         Left            =   360
         Picture         =   "frmMain.frx":6D1C
         Top             =   315
         Width           =   390
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Press 'start' to begin!"
         Height          =   195
         Left            =   1080
         TabIndex        =   40
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ready to start"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   840
         TabIndex        =   38
         Top             =   360
         Width           =   1755
      End
      Begin VB.Label Label28 
         Caption         =   "The wizard has now collected enough information to begin the image extracting process."
         Height          =   495
         Left            =   1080
         TabIndex        =   37
         Top             =   840
         Width           =   4335
      End
   End
   Begin VB.Frame picFrameFinished 
      BackColor       =   &H00FFFFFF&
      Height          =   4455
      Left            =   2160
      TabIndex        =   29
      Top             =   -240
      Width           =   4815
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   240
         Picture         =   "frmMain.frx":7265
         ScaleHeight     =   480
         ScaleWidth      =   420
         TabIndex        =   30
         Top             =   3180
         Width           =   420
      End
      Begin VB.TextBox txtBatchLog 
         Height          =   1695
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   56
         Text            =   "frmMain.frx":7758
         Top             =   1320
         Width           =   3855
         Visible         =   0   'False
      End
      Begin VB.Image imgShowLog 
         Height          =   240
         Left            =   375
         Picture         =   "frmMain.frx":7764
         Top             =   1410
         Width           =   240
         Visible         =   0   'False
      End
      Begin VB.Label lblLinkButton 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Show log..."
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   5
         Left            =   675
         TabIndex        =   57
         Top             =   1440
         Width           =   795
         Visible         =   0   'False
      End
      Begin VB.Image imgDestinationFolder 
         Height          =   225
         Left            =   360
         Picture         =   "frmMain.frx":7B3C
         Top             =   1410
         Width           =   255
      End
      Begin VB.Label lblLinkButton 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Click here to open destination folder..."
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   3
         Left            =   675
         TabIndex        =   35
         Top             =   1440
         Width           =   2670
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Finished!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   34
         Top             =   360
         Width           =   1170
      End
      Begin VB.Label lblExecuteInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "## images extracted!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   32
         Top             =   960
         Width           =   3615
      End
      Begin VB.Label lblLinkButton 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Download PDF Image Extraction Wizard"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   31
         Top             =   3780
         Width           =   3690
      End
      Begin VB.Label lblFlashAd1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tip: If you need to extract images from PDF documents, use PDF Wiz instead:"
         Height          =   615
         Left            =   840
         TabIndex        =   33
         Top             =   3300
         Width           =   3255
      End
   End
   Begin VB.Frame picFrameBatchInput 
      Height          =   3375
      Left            =   0
      TabIndex        =   41
      Top             =   840
      Width           =   6855
      Begin VB.PictureBox picTooltip 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   0
         Left            =   6255
         Picture         =   "frmMain.frx":7F34
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   61
         Top             =   1920
         Width           =   225
      End
      Begin VB.CheckBox chkBatchCreateFolders 
         Caption         =   "Create a folder for each document"
         Height          =   255
         Left            =   3480
         TabIndex        =   52
         Top             =   1920
         Width           =   3015
      End
      Begin ComctlLib.ListView lwFiles 
         Height          =   1575
         Left            =   240
         TabIndex        =   43
         ToolTipText     =   "Drop files or folders here..."
         Top             =   480
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   2778
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         OLEDropMode     =   1
         _Version        =   327682
         SmallIcons      =   "ImageList"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         OLEDropMode     =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   ""
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox txtBatchOutput 
         Height          =   285
         Left            =   3765
         OLEDropMode     =   1  'Manual
         TabIndex        =   49
         Top             =   960
         Width           =   2370
      End
      Begin VB.OptionButton optBatchOutput 
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   48
         Top             =   960
         Width           =   495
      End
      Begin VB.OptionButton optBatchOutput 
         Caption         =   "Same as each file's input folder"
         Height          =   255
         Index           =   0
         Left            =   3480
         TabIndex        =   46
         Top             =   600
         Value           =   -1  'True
         Width           =   2895
      End
      Begin VB.CommandButton cmdBatchOutput 
         Height          =   375
         Left            =   6240
         Picture         =   "frmMain.frx":8301
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Browse for a folder to save the extracted images in..."
         Top             =   915
         Width           =   375
      End
      Begin VB.ListBox lstFilesComplete 
         Height          =   1425
         Left            =   3600
         TabIndex        =   42
         Top             =   3000
         Width           =   1695
         Visible         =   0   'False
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add Files"
         Height          =   750
         Left            =   240
         Picture         =   "frmMain.frx":86F9
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Add documents to the list. To add all files in a folder, drop the folder onto the list."
         Top             =   2160
         Width           =   1575
      End
      Begin VB.CommandButton cmdRemoveSelected 
         Caption         =   "Remove"
         Height          =   750
         Left            =   2040
         Picture         =   "frmMain.frx":8B65
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Remove the selected items from the list, or all items if none is selected."
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label lblDocsToProcess 
         AutoSize        =   -1  'True
         Caption         =   "Documents to process:"
         Height          =   195
         Left            =   240
         TabIndex        =   59
         Top             =   240
         Width           =   1635
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Options:"
         Height          =   195
         Left            =   3360
         TabIndex        =   58
         Top             =   1560
         Width           =   585
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Output folder:"
         Height          =   195
         Left            =   3360
         OLEDropMode     =   1  'Manual
         TabIndex        =   47
         Top             =   240
         Width           =   960
      End
      Begin ComctlLib.ImageList ImageList 
         Left            =   5520
         Top             =   3000
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   327682
      End
      Begin VB.Image imgListviewIcon 
         Height          =   210
         Left            =   5300
         Picture         =   "frmMain.frx":8FB8
         Top             =   3000
         Width           =   180
         Visible         =   0   'False
      End
   End
   Begin VB.Frame picFrameInput 
      Height          =   3375
      Left            =   0
      OLEDropMode     =   1  'Manual
      TabIndex        =   26
      Top             =   1080
      Width           =   6855
      Begin VB.PictureBox picTooltip 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   1
         Left            =   2235
         Picture         =   "frmMain.frx":9093
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   60
         Top             =   2310
         Width           =   225
      End
      Begin VB.CheckBox chkCreateFolders 
         Caption         =   "Create a folder here"
         Height          =   255
         Left            =   480
         OLEDropMode     =   1  'Manual
         TabIndex        =   5
         ToolTipText     =   "Select this to create a folder to save extracted files in"
         Top             =   2310
         Width           =   2175
      End
      Begin VB.CommandButton cmdOutput 
         Height          =   375
         Left            =   6120
         OLEDropMode     =   1  'Manual
         Picture         =   "frmMain.frx":9460
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Browse for a folder to save the extracted images in..."
         Top             =   1845
         Width           =   375
      End
      Begin VB.CommandButton cmdInput 
         Height          =   375
         Left            =   6120
         OLEDropMode     =   1  'Manual
         Picture         =   "frmMain.frx":9858
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Browse to select the document file you want to extract images from..."
         Top             =   375
         Width           =   375
      End
      Begin VB.TextBox txtOutput 
         Height          =   285
         Left            =   360
         OLEDropMode     =   1  'Manual
         TabIndex        =   3
         ToolTipText     =   "This is the folder where all extracted images are saved."
         Top             =   1905
         Width           =   5655
      End
      Begin VB.TextBox txtInput 
         Height          =   285
         Left            =   360
         OLEDropMode     =   1  'Manual
         TabIndex        =   1
         ToolTipText     =   "This is the document you want to extract images from. Most common document formats are supported."
         Top             =   435
         Width           =   5655
      End
      Begin VB.Label Label33 
         Caption         =   $"frmMain.frx":9C50
         ForeColor       =   &H80000011&
         Height          =   915
         Left            =   360
         OLEDropMode     =   1  'Manual
         TabIndex        =   63
         Top             =   720
         Width           =   5475
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Output folder:"
         Height          =   195
         Left            =   360
         OLEDropMode     =   1  'Manual
         TabIndex        =   28
         Top             =   1665
         Width           =   960
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Document:"
         Height          =   195
         Left            =   360
         OLEDropMode     =   1  'Manual
         TabIndex        =   27
         Top             =   195
         Width           =   780
      End
   End
   Begin VB.Frame picFrameExecute 
      Height          =   3015
      Left            =   600
      TabIndex        =   18
      Top             =   1080
      Width           =   6255
      Begin VB.PictureBox picProgressWorking 
         AutoSize        =   -1  'True
         Height          =   300
         Left            =   360
         Picture         =   "frmMain.frx":9D6B
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   2640
         Width           =   300
         Visible         =   0   'False
      End
      Begin VB.PictureBox picProgressFinished 
         AutoSize        =   -1  'True
         Height          =   300
         Left            =   720
         Picture         =   "frmMain.frx":A0CE
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   2640
         Width           =   300
         Visible         =   0   'False
      End
      Begin VB.PictureBox picProgress3 
         AutoSize        =   -1  'True
         Height          =   300
         Left            =   1440
         Picture         =   "frmMain.frx":A421
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   2430
         Width           =   300
         Visible         =   0   'False
      End
      Begin VB.PictureBox picProgress2 
         AutoSize        =   -1  'True
         Height          =   300
         Left            =   1440
         Picture         =   "frmMain.frx":A774
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1830
         Width           =   300
         Visible         =   0   'False
      End
      Begin VB.PictureBox picProgress1 
         AutoSize        =   -1  'True
         Height          =   300
         Left            =   1440
         Picture         =   "frmMain.frx":AAC7
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1230
         Width           =   300
         Visible         =   0   'False
      End
      Begin VB.Label lblExtractingInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Processing, please wait..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1320
         TabIndex        =   50
         Top             =   360
         Width           =   3225
      End
      Begin VB.Image Image1 
         Height          =   795
         Left            =   240
         Picture         =   "frmMain.frx":AE1A
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Please wait while the wizard is extracting your images."
         Height          =   195
         Left            =   1320
         TabIndex        =   25
         Top             =   720
         Width           =   3795
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Moving files to destination folder"
         Height          =   195
         Left            =   1800
         TabIndex        =   24
         Top             =   2460
         Width           =   2265
         Visible         =   0   'False
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Processing images"
         Height          =   195
         Left            =   1800
         TabIndex        =   23
         Top             =   1860
         Width           =   1320
         Visible         =   0   'False
      End
      Begin VB.Label lblExtractingData 
         AutoSize        =   -1  'True
         Caption         =   "Extracting data"
         Height          =   195
         Left            =   1800
         TabIndex        =   22
         Top             =   1260
         Width           =   1065
         Visible         =   0   'False
      End
   End
   Begin VB.ComboBox cmbBatchFiles 
      Height          =   315
      ItemData        =   "frmMain.frx":B722
      Left            =   720
      List            =   "frmMain.frx":B729
      Style           =   2  'Dropdown List
      TabIndex        =   55
      Top             =   4560
      Width           =   1815
      Visible         =   0   'False
   End
   Begin VB.CheckBox chkBatchMode 
      Caption         =   "Batch Mode"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      ToolTipText     =   "Batch mode allows you to process multiple documents at once."
      Top             =   4620
      Width           =   1455
   End
   Begin VB.Timer tmrRestoreLinkLabels 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2400
      Top             =   4680
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   6240
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   6240
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   0
      X2              =   6240
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   0
      X2              =   6240
      Y1              =   4440
      Y2              =   4440
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const FRAME_WELCOME = 1
Private Const FRAME_INPUT = 2
Private Const FRAME_BEGIN = 3
Private Const FRAME_EXECUTE = 4
Private Const FRAME_FINISHED = 5
Private Const FRAME_BATCH_INPUT = 6

Private currentFrame As Integer
Private tempFolder As String
Private sLastDir As String
Private txtBasename As String

Dim myFso
Dim bHaveFileSystemObject As Boolean
Dim myTT() As CTooltip
Private WithEvents m_cUnzip As cUnzip
Attribute m_cUnzip.VB_VarHelpID = -1

Private Sub Form_Load()

    On Error Resume Next

    ChDrive Left(App.Path, 1)
    ChDir App.Path

    isRelease = True    'set to false if the next line executes
    Debug.Assert Check_If_Release() 'only executes if run from within ide

    Call InitPortable

    frmMain.Caption = frmMain.Caption & " 4.1"

    currentFrame = FRAME_WELCOME
    updateFrames
    
    'Get temp path (no trailing slash)
    Call getSystemFolders(sysFolders)
    tempFolder = sysFolders.Temp

    SetIcon Me.hWnd, "AAA", True

    'Resize Controls (everything originates from size of picSplash & picFrameTop)
    Line1.Y1 = picFrameTop.Height + 15
    Line2.Y1 = picFrameTop.Height + 15
    Line3.Y1 = cmdNext.Top - 210
    Line4.Y1 = cmdNext.Top - 210

    Line1.Y2 = Line1.Y1
    Line2.Y2 = Line2.Y1
    Line3.Y2 = Line3.Y1
    Line4.Y2 = Line4.Y1

    Line1.X1 = 0
    Line2.X1 = 0
    Line3.X1 = 0
    Line4.X1 = 0
    
    Line1.X2 = picFrameTop.Width
    Line2.X2 = picFrameTop.Width
    Line3.X2 = picFrameTop.Width
    Line4.X2 = picFrameTop.Width

    picFrameWelcome.BorderStyle = 0
    picFrameWelcome.Top = 0
    picFrameWelcome.Height = cmdNext.Top - 225
    picFrameWelcome.Left = picSplash.Width
    picFrameWelcome.Width = picFrameTop.Width - picSplash.Width

    picFrameInput.BorderStyle = 0
    picFrameInput.Top = picFrameTop.Height + 45
    picFrameInput.Height = cmdNext.Top - 225 - picFrameTop.Height - 45
    picFrameInput.Left = 0
    picFrameInput.Width = picFrameTop.Width

    picFrameBegin.BorderStyle = 0
    picFrameBegin.Top = picFrameInput.Top
    picFrameBegin.Height = picFrameInput.Height
    picFrameBegin.Left = picFrameInput.Left
    picFrameBegin.Width = picFrameInput.Width

    picFrameExecute.BorderStyle = 0
    picFrameExecute.Top = picFrameInput.Top
    picFrameExecute.Height = picFrameInput.Height
    picFrameExecute.Left = picFrameInput.Left
    picFrameExecute.Width = picFrameInput.Width

    picFrameFinished.BorderStyle = 0
    picFrameFinished.Top = 0
    picFrameFinished.Height = cmdNext.Top - 225
    picFrameFinished.Left = picSplash.Width
    picFrameFinished.Width = picFrameTop.Width - picSplash.Width

    picFrameBatchInput.BorderStyle = 0
    picFrameBatchInput.Top = picFrameTop.Height + 45
    picFrameBatchInput.Height = cmdNext.Top - 225 - picFrameTop.Height - 45
    picFrameBatchInput.Left = 0
    picFrameBatchInput.Width = picFrameTop.Width

    If IsThemed() Then
        FixThemeSupport Controls
    End If

    lwFiles.ColumnHeaders.Item(1).Width = lwFiles.Width - 600
    Call ImageList.ListImages.Add(, , imgListviewIcon)

    Dim tmp As Integer
    chkSkipReadyPage.Value = GetSettingEX("RL Vision", "Office Wiz", "chkSkipReadyPage", 0)
    bFontSizeWarningShown = GetSettingEX("RL Vision", "Office Wiz", "bFontSizeWarningShown", False)
    chkBatchCreateFolders = GetSettingEX("RL Vision", "Office Wiz", "chkBatchCreateFolders", 0)

    txtBatchOutput = GetSettingEX("RL Vision", "Office Wiz", "txtBatchOutput", "")
    tmp = GetSettingEX("RL Vision", "Office Wiz", "optBatchOutput", 0)
    optBatchOutput(tmp).Value = True
    
    chkCreateFolders = GetSettingEX("RL Vision", "Office Wiz", "chkCreateFolders", 0)

    If Dir(txtBatchOutput, vbDirectory) = "" Then
        txtBatchOutput = ""
        optBatchOutput(0).Value = True
    End If
   
    chkBatchMode = GetSettingEX("RL Vision", "Office Wiz", "chkBatchMode", 0)
   
    Dim sTmp As String
    sTmp = GetSettingEX("RL Vision", "Office Wiz", "txtInput", "")
    If sTmp <> "" And Dir(sTmp, vbArchive + vbHidden + vbNormal + vbReadOnly + vbSystem) <> "" Then
        txtInput = GetSettingEX("RL Vision", "Office Wiz", "txtInput", "")
        txtOutput = GetSettingEX("RL Vision", "Office Wiz", "txtOutput", "")
        txtBasename = GetSettingEX("RL Vision", "Office Wiz", "txtBasename", "")
        Call SetInputFile(txtInput, False)
    End If

    
    If Screen.TwipsPerPixelX <> 15 Or Screen.TwipsPerPixelY <> 15 Then
        lblLinkButton(2).Left = lblLinkButton(2).Left + 800
    End If
    
   'Init FileSystemObject
   On Error Resume Next
   Err.Clear
   bHaveFileSystemObject = True
   Set myFso = CreateObject("Scripting.FileSystemObject")
   If Err.Number <> 0 Then
       bHaveFileSystemObject = False
   End If
   On Error GoTo 0


    'setup tooltips
    ReDim myTT(picTooltip.Count)
    Dim n As Integer
    For n = 0 To picTooltip.Count - 1
        Set myTT(n) = New CTooltip
        myTT(n).Style = TTBalloon
        myTT(n).Icon = TTIconInfo
        myTT(n).VisibleTime = 30000
        myTT(n).DelayTime = 200

        Select Case n
        Case 0:
            myTT(n).Title = "Create folder"
            myTT(n).TipText = "If this is selected, a folder is created in the output folder where extracted images are saved. The new folder is named after the filename of the input document."
        Case 1:
            myTT(n).Title = "Create folder for each document"    'batch
            myTT(n).TipText = "If this is selected, extracted images will be saved in separate folders for each document. The folder is created in the output folder and named after the original document."
        End Select
        myTT(n).Create picTooltip(n).hWnd
    Next

End Sub

Private Sub Form_Activate()

    If cmdNext.Enabled = True Then cmdNext.SetFocus

    'command line
    If Command <> "" Then
        Dim sFil As String
        sFil = Trim(Command)
        If Left(sFil, 1) = Chr(34) Then sFil = Mid(sFil, 2)
        If Right(sFil, 1) = Chr(34) Then sFil = Mid(sFil, 1, Len(sFil) - 1)
        If Dir(sFil, vbArchive + vbHidden + vbHidden + vbReadOnly + vbSystem) <> "" Then
            SetInputFile (sFil)
            Call cmdNext_Click
        End If
    End If

    If IsWin7Plus Then
        If Screen.TwipsPerPixelX = 15 And UCase(Left(GetWin7Font(), 11)) = "SSERIFF.FON" Then
            If bFontSizeWarningShown = False Then
                Call MsgBox("Because of a bug in Windows 7, your system is using a font set that is too large for your current DPI settings. This will cause PDF Wiz (and other programs) to look distorted. It will still work, but I strongly recommend you to fix it. A page will now open with help on how to solve the problem!", vbExclamation, "Important Information")
                Call ShellExecute(Me.hWnd, vbNullString, "http://www.rlvision.com/misc/windows_7_font_bug.asp", vbNullString, "c:\", SW_NORMAL)
                bFontSizeWarningShown = True
            End If
        End If
    End If

End Sub

Private Sub updateFrames()

    picSplash.Visible = False
    picFrameTop.Visible = False
    picFrameWelcome.Visible = False
    picFrameInput.Visible = False
    picFrameExecute.Visible = False
    picFrameFinished.Visible = False
    picFrameBegin.Visible = False
    picFrameBatchInput.Visible = False

    Line1.Visible = True
    Line2.Visible = True
    Line3.Visible = True
    Line4.Visible = True
    
    chkBatchMode.Visible = False
    cmbBatchFiles.Visible = False
    
    Select Case currentFrame
    
        Case FRAME_WELCOME
        
            picSplash.Visible = True
            picFrameWelcome.Visible = True
            
            cmdBack.Enabled = False
            cmdNext.Enabled = True
            cmdCancel.Enabled = True
            cmdNext.Caption = "Next >"
    
            Line1.Visible = False
            Line2.Visible = False

        Case FRAME_INPUT
        
            picFrameTop.Visible = True
            picFrameInput.Visible = True
    
            cmdNext.Enabled = False
            cmdBack.Enabled = True
            cmdCancel.Enabled = True
            cmdCancel.Caption = "Exit"

            chkBatchMode.Visible = True
            
            Call validateFrameInput 'enable/disable cmdNext
                    
            lblTopTitle = "Input && Output"
            lblTopText = "What document do you want to extract images from, and where should the wizard place the extracted image files?"
        
            cmdNext.Caption = "Next >"
            If cmdNext.Enabled = True Then cmdNext.SetFocus
            
            cmdInput.Refresh
            cmdOutput.Refresh


        Case FRAME_BATCH_INPUT

            picFrameTop.Visible = True
            picFrameBatchInput.Visible = True
    
            cmdNext.Enabled = False
            cmdBack.Enabled = True
            cmdCancel.Enabled = True
            cmdCancel.Caption = "Exit"
            
            chkBatchMode.Visible = True
            
            Call validateFrameInput 'enable/disable cmdNext
                    
            lblTopTitle = "Batch Input && Output"
            lblTopText = "Add one or more document to the batch processing list. Use the add button, or drop files and folders onto the list."

            cmdNext.Caption = "Next >"
            If cmdNext.Enabled = True Then cmdNext.SetFocus Else cmdAdd.SetFocus


            cmdAdd.Refresh
            cmdRemoveSelected.Refresh
            cmdBatchOutput.Refresh


        Case FRAME_BEGIN
            picFrameTop.Visible = True
            picFrameBegin.Visible = True
            
            lblTopTitle = "Ready to Start"
            lblTopText = "The wizard has now collected enough information to begin extracting the images!"

            cmdNext.Enabled = True
            cmdBack.Enabled = True
            cmdCancel.Enabled = True
            cmdNext.Caption = "Start >"
        
        Case FRAME_EXECUTE
        
            cmdNext.Enabled = False
            cmdBack.Enabled = False
            cmdCancel.Enabled = False
            
            picFrameTop.Visible = True
            picFrameExecute.Visible = True
            
            lblTopTitle = "Extracting Images"
            lblTopText = "Please wait while images are being extracted and processed."


        Case FRAME_FINISHED

            picSplash.Visible = True
            picFrameFinished.Visible = True
   
            Line1.Visible = False
            Line2.Visible = False
            
            cmdNext.Enabled = False
            cmdBack.Enabled = True
            cmdCancel.Enabled = True
            
            cmdCancel.Caption = "Close"
            cmdCancel.SetFocus
    
            lblTopTitle = "Finished"
            lblTopText = "The wizard has finished extracting images! Press the exit button to quit, or the back button to process another file."
            
            DoEvents
    
    End Select

End Sub

Private Sub validateFrameInput()
    
    If currentFrame = FRAME_INPUT Then
        If chkBatchMode.Value = 1 Then
            currentFrame = FRAME_BATCH_INPUT
        End If
    End If
    
    If currentFrame <> FRAME_INPUT And currentFrame <> FRAME_BATCH_INPUT Then Exit Sub
    
    On Error Resume Next
    
    If currentFrame = FRAME_INPUT Then
        If txtInput = "" Then GoTo fail
        If txtOutput = "" Then GoTo fail
        
        If Dir(txtInput, vbArchive + vbHidden + vbReadOnly + vbSystem) = "" Then GoTo fail
        Dim d As Integer
        d = GetAttr(txtInput)
        If (d And vbDirectory) = vbDirectory Then GoTo fail
        
        cmdNext.Enabled = True
        Exit Sub
    End If
    
    If currentFrame = FRAME_BATCH_INPUT Then
        If lwFiles.ListItems.Count = 0 Then GoTo fail
        If optBatchOutput(1).Value = True And txtBatchOutput = "" Then GoTo fail
        cmdNext.Enabled = True
        Exit Sub
    End If
    
fail:
    cmdNext.Enabled = False

End Sub

Private Sub cmdBack_Click()

    If currentFrame = FRAME_BATCH_INPUT Then currentFrame = FRAME_INPUT
    
    If currentFrame = FRAME_BEGIN Or currentFrame = FRAME_FINISHED Then
        currentFrame = FRAME_INPUT
    Else
        currentFrame = currentFrame - 1
    End If
    
    
    If currentFrame = FRAME_INPUT Then
        If chkBatchMode.Value = 1 Then currentFrame = FRAME_BATCH_INPUT
    End If

    updateFrames

End Sub


Private Sub cmdNext_Click()

    On Error GoTo errHandler
    
    Dim n As Integer
    
    If currentFrame = FRAME_INPUT Then
        ReDim documentOptions(1 To 1)
        iCurrentDocument = 1
        documentOptions(1).txtInput = txtInput
        documentOptions(1).txtOutput = txtOutput
        documentOptions(1).txtBasename = txtBasename
        documentOptions(1).iCreateFolder = chkCreateFolders.Value
                
    ElseIf currentFrame = FRAME_BATCH_INPUT Then
    
        cmbBatchFiles.Clear
        cmbBatchFiles.AddItem "All Documents"
        
        ReDim documentOptions(1 To lwFiles.ListItems.Count)
        For n = 1 To lwFiles.ListItems.Count
            documentOptions(n).txtInput = lstFilesComplete.List(n - 1)
            If optBatchOutput(0).Value = True Then
                documentOptions(n).txtOutput = GetPathFromFilename(lstFilesComplete.List(n - 1))
            ElseIf optBatchOutput(1).Value = True Then
                documentOptions(n).txtOutput = txtBatchOutput
            End If
            documentOptions(n).txtBasename = lwFiles.ListItems.Item(n).Text
            
            documentOptions(n).iCreateFolder = chkBatchCreateFolders
            
            cmbBatchFiles.AddItem lwFiles.ListItems.Item(n).Text
        Next
        
        cmbBatchFiles.ListIndex = 0

        iCurrentDocument = -1 'all

    End If

    If currentFrame = FRAME_BATCH_INPUT Then currentFrame = FRAME_INPUT
    
    currentFrame = currentFrame + 1
    
    If currentFrame = FRAME_INPUT Then
        If chkBatchMode.Value = 1 Then currentFrame = FRAME_BATCH_INPUT
    End If
    
    If currentFrame = FRAME_BEGIN And chkSkipReadyPage.Value = 1 Then
        currentFrame = currentFrame + 1
    End If
    

    updateFrames
    
    If currentFrame = FRAME_EXECUTE Then
    
        imgDestinationFolder.Visible = True
        lblLinkButton(3).Visible = True
        
        'extract
        For iCurrentDocument = LBound(documentOptions) To UBound(documentOptions)
            picProgress1.Picture = picProgressWorking.Picture
            If LBound(documentOptions) <> UBound(documentOptions) Then
                lblExtractingInfo = "Processing document " & iCurrentDocument & " of " & UBound(documentOptions) & " ..."
                DoEvents
            End If
            Call doExtraction
        Next
        
        
        'hide open destination folder link
        If chkBatchMode.Value = 1 Then
            lblLinkButton(3).Visible = False
            imgDestinationFolder.Visible = False
        End If

        
        'show output info
        If LBound(documentOptions) = UBound(documentOptions) Then
            lblExecuteInfo = documentOptions(LBound(documentOptions)).sExecuteInfo
            txtBatchLog.Visible = False
            lblLinkButton(5).Visible = False
            imgShowLog.Visible = False
        Else
            txtBatchLog = ""
            txtBatchLog.Visible = False
            lblLinkButton(5).Visible = True
            imgShowLog.Visible = True
            Dim c As Integer: Dim e As Integer: Dim i As Integer
            c = 0: e = 0: i = 0

            For iCurrentDocument = LBound(documentOptions) To UBound(documentOptions)
                c = c + documentOptions(iCurrentDocument).iImagesExtracted
                If documentOptions(iCurrentDocument).iImagesExtracted > 0 Then i = i + 1
                If documentOptions(iCurrentDocument).bError = True Then e = e + 1
                txtBatchLog = txtBatchLog & "[" & documentOptions(iCurrentDocument).txtBasename & "] " & documentOptions(iCurrentDocument).sExecuteInfo & vbNewLine
            Next

            If i > 0 Then lblExecuteInfo = c & " images extracted from " & i & " documents!" Else lblExecuteInfo = "No images extracted..."
            If e > 0 Then lblExecuteInfo = lblExecuteInfo & " " & e & " errors occurred..."
            
        End If


        currentFrame = currentFrame + 1
        updateFrames
    End If


    Exit Sub
errHandler:
    Call MsgBox("Error " & Err.Number & ": " & Err.Description, vbCritical)

End Sub

Private Sub cmdCancel_Click()

    If currentFrame <> FRAME_FINISHED Then
        Unload Me
    Else
        Unload Me
    End If

End Sub

Private Sub cmdInput_Click()

    ' api open dialog box '''''''''
    Dim tOPENFILENAME As OPENFILENAME
    Dim lResult As Long
    Dim vFiles As Variant
    Dim lIndex As Long, lStart As Long
    Dim sAddFolder As String

    
    Dim sFile As String, InitDir As String
    If txtInput = "" Then
        sFile = ""
        InitDir = sysFolders.MyDocuments
    Else
        sFile = txtInput
        InitDir = ""
    End If

    cmdInput.Refresh
    cmdOutput.Refresh


    With tOPENFILENAME
        .Flags = OFN_EXPLORER Or OFN_FILEMUSTEXIST Or OFN_LONGNAMES
        .hwndOwner = hWnd
        .nMaxFile = 32768
        .lpstrFilter = "Supported Documents" & Chr(0) & "*.doc;*.dot;*.ppt;*.pot;*.pps;*.docm;*.docx;*.dotm;*.dotx;*.potm;*.potx;*.pptm;*.pptx;*.ppsm;*.ppsx;*.sldm;*.xlsm;*.xlsx;*.xltm;*.xltx;*.odt;*.odp;*.ods;*.ott;*.odg;*.ots;*.otp;*.otg;*.pages;*.numbers;*.template;*.sxw;*.stw;*.sxc;*.stc;*.sxi;*.sti;*.epub;*.cbz;*.fb2;*.xps;*.oxps;*.dwfx;*.chm;*.swf" & Chr(0) & "All Files" & Chr(0) & "*.*" & Chr(0) & Chr(0)
        .lpstrFile = Space(.nMaxFile - 1) & Chr(0)
        .lpstrInitialDir = InitDir
        .lStructSize = Len(tOPENFILENAME)
    End With

    lResult = GetOpenFileName(tOPENFILENAME)

    If lResult > 0 Then
        
        Screen.MousePointer = 11
        
        Dim sTmp As String
        sTmp = tOPENFILENAME.lpstrFile
        If InStr(sTmp, Chr(0)) > 0 Then sTmp = Mid(sTmp, 1, InStr(sTmp, Chr(0)) - 1)
        SetInputFile (sTmp)
    
        Call validateFrameInput 'enable/disable cmdNext
        Screen.MousePointer = 0
    
    End If

End Sub

Private Sub cmdOutput_Click()

    Dim sFolder As String
    
    If Right(txtOutput, 1) = "\" And Right(txtOutput, 2) <> ":\" Then txtOutput = Left(txtOutput, Len(txtOutput) - 1)    'stip right slash if present
    sFolder = txtOutput
    sFolder = BrowseForFolder(Me, "Select output folder:", , , sFolder, , , , , sysFolders.NetHood)

    cmdInput.Refresh
    cmdOutput.Refresh

    If sFolder <> "" Then txtOutput = sFolder

End Sub


Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If currentFrame <> FRAME_WELCOME And currentFrame <> FRAME_INPUT Then Exit Sub
    Call DoDrop(Data)
    If currentFrame = FRAME_WELCOME Then Call cmdNext_Click
End Sub

Private Sub imgOrder_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DoDrop(Data)
    Call cmdNext_Click
End Sub

Private Sub Label16_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DoDrop(Data)
End Sub

Private Sub Label17_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DoDrop(Data)
End Sub

Private Sub Label32_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DoDrop(Data)
    Call cmdNext_Click
End Sub

Private Sub Label33_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DoDrop(Data)
End Sub

Private Sub Label4_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DoDrop(Data)
End Sub

Private Sub Label8_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DoDrop(Data)
End Sub

Private Sub lblTopText_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If currentFrame = FRAME_INPUT Then Call DoDrop(Data)
End Sub

Private Sub lblTopTitle_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If currentFrame = FRAME_INPUT Then Call DoDrop(Data)
End Sub

Private Sub lwFiles_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDelete Then
        Dim n As Integer
    
restart:
        For n = 1 To lwFiles.ListItems.Count
            If lwFiles.ListItems(n).Selected = True Then
                lwFiles.ListItems.Remove (n)
                lstFilesComplete.RemoveItem (n - 1)
                GoTo restart
            End If
        Next
    
        Call validateFrameInput
    End If

    lblDocsToProcess = "Documents to process (" & lstFilesComplete.ListCount & "):"

End Sub

Private Sub lwFiles_OLEDragDrop(Data As ComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Screen.MousePointer = 11

    If Data.GetFormat(15) = True Then

        Dim n As Integer
        Dim sTmp As String
        Dim itmX As ListItem
        Dim test1 As String
        Dim test2 As String
        
        For n = 1 To Data.Files.Count
        
            If (GetAttr(Data.Files(n)) And vbDirectory) = vbDirectory Then
            
                ' add files in folder
                Dim myList As New clsListEmu
                Call RecursiveGetFolderContent(Data.Files(n), myList, True, False, True)
                Dim i As Long
                For i = 0 To myList.ListCount - 1
                    test1 = Right(LCase(myList.List(i)), 4)
                    test2 = ".doc .dot .ppt .pot .pps .docm .docx .dotm .dotx .potm .potx .pptm .pptx .ppsm .ppsx .sldm .xlsm .xlsx .xltm .xltx .odt .odp .ods .ott .odg .ots .otp .otg .pages .numbers .template .sxw .stw .sxc .stc .sxi .sti .epub .cbz .fb2 .xps .oxps .dwfx .chm .swf"
                    If InStr(test2, test1) > 0 Then
                        AddFile (myList.List(i))
                    End If
                Next
                
            Else
                'add single file
                test1 = Right(LCase(Data.Files(n)), 4)
                test2 = ".doc .dot .ppt .pot .pps .docm .docx .dotm .dotx .potm .potx .pptm .pptx .ppsm .ppsx .sldm .xlsm .xlsx .xltm .xltx .odt .odp .ods .ott .odg .ots .otp .otg .pages .numbers .template .sxw .stw .sxc .stc .sxi .sti .epub .cbz .fb2 .xps .oxps .dwfx .chm .swf"
                If InStr(test2, test1) > 0 Then
                    If AddFile(Data.Files(n)) = False Then
                        Exit For
                    End If
                End If
            End If

        Next
        
    End If

    Call validateFrameInput
    
    Screen.MousePointer = 0

End Sub

Private Sub optBatchOutput_Click(Index As Integer)
    Call validateFrameInput 'enable/disable cmdNext
End Sub

Private Sub picFrameTop_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If currentFrame = FRAME_INPUT Then Call DoDrop(Data)
End Sub

Private Sub txtBatchOutput_Change()
    optBatchOutput(1).Value = True
    Call validateFrameInput 'enable/disable cmdNext
End Sub

Private Sub txtBatchOutput_GotFocus()
    optBatchOutput(1).Value = True
    Call validateFrameInput 'enable/disable cmdNext
End Sub

Private Sub txtInput_Change()
    Call validateFrameInput 'enable/disable cmdNext
End Sub

Private Sub txtInput_GotFocus()
    txtInput.Tag = txtInput.Text
End Sub

Private Sub txtInput_LostFocus()
    If txtInput.Tag <> txtInput.Text Then
        Call SetInputFile(txtInput)
    End If
End Sub

Private Sub txtInput_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DoDrop(Data)
End Sub

Private Sub txtOutput_Change()
    Call validateFrameInput 'enable/disable cmdNext
End Sub

Private Sub txtOutput_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DoDrop(Data)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    
    SaveSettingEX "RL Vision", "Office Wiz", "txtInput", txtInput
    SaveSettingEX "RL Vision", "Office Wiz", "txtOutput", txtOutput
    SaveSettingEX "RL Vision", "Office Wiz", "txtBasename", txtBasename
    SaveSettingEX "RL Vision", "Office Wiz", "chkSkipReadyPage", chkSkipReadyPage.Value
    SaveSettingEX "RL Vision", "Office Wiz", "bFontSizeWarningShown", bFontSizeWarningShown
    SaveSettingEX "RL Vision", "Office Wiz", "chkBatchMode", chkBatchMode
    SaveSettingEX "RL Vision", "Office Wiz", "chkBatchCreateFolders", chkBatchCreateFolders
    SaveSettingEX "RL Vision", "Office Wiz", "chkCreateFolders", chkCreateFolders
    
    SaveSettingEX "RL Vision", "Office Wiz", "txtBatchOutput", txtBatchOutput
    Dim tmp As Integer
    If optBatchOutput(0).Value = True Then tmp = 0
    If optBatchOutput(1).Value = True Then tmp = 1
    SaveSettingEX "RL Vision", "Office Wiz", "optBatchOutput", tmp

    Call UnloadXpApp

End Sub

Private Sub picFrameInput_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DoDrop(Data)
End Sub

Private Sub SetInputFile(sFile As String, Optional bUpdateOutput As Boolean = True)

    If Dir(sFile) = "" Or InStrRev(sFile, "\") = 0 Or InStrRev(sFile, ".") = 0 Then
        Exit Sub
    End If
    
    sFile = Trim(sFile)
    txtInput = sFile
    If bUpdateOutput = True Then txtBasename = Mid(sFile, InStrRev(sFile, "\") + 1, InStrRev(sFile, ".") - InStrRev(sFile, "\") - 1)
    If bUpdateOutput = True Then txtOutput = GetPathFromFilename(sFile)

Exit Sub

    ChDrive (App.Path)
    ChDir (App.Path)
    
End Sub

Private Sub DoDrop(Data As DataObject)

    If Data.GetFormat(15) = True Then

        If Data.Files.Count > 1 Then
            Call MsgBox("You can only drop one file at a time...", vbExclamation)
        End If
        If Data.Files.Count = 1 Then
            SetInputFile (Data.Files(1))
        End If
    
    End If

End Sub

Private Sub picSplash_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DoDrop(Data)
    Call cmdNext_Click
End Sub
Private Sub picFrameWelcome_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DoDrop(Data)
    Call cmdNext_Click
End Sub
Private Sub Label3_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DoDrop(Data)
    Call cmdNext_Click
End Sub
Private Sub Label5_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DoDrop(Data)
    Call cmdNext_Click
End Sub
Private Sub Label7_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DoDrop(Data)
    Call cmdNext_Click
End Sub

Private Sub lblLinkButton_Click(Index As Integer)

    Dim bIsStandard As Boolean: Dim bIsPro As Boolean
    
    Select Case Index
    
    Case 0:
    
    Case 1:
    
    Case 2:
        'open website
        Call ShellExecute(Me.hWnd, "open", "http://www.rlvision.com/script/redirect.asp?app=officewiz", vbNullString, vbNullString, SW_NORMAL)
        
    Case 3:
        'open destination folder
        If chkBatchMode.Value = 1 Then
            If optBatchOutput(1).Value = True Then
                Call ShellExecute(Me.hWnd, "open", txtBatchOutput, vbNullString, vbNullString, SW_NORMAL)
            End If
        Else
            Call ShellExecute(Me.hWnd, "open", txtOutput, vbNullString, vbNullString, SW_NORMAL)
        End If

    Case 4:
        'pdf wiz
        Call ShellExecute(Me.hWnd, "open", "http://www.rlvision.com/downloads.asp", vbNullString, vbNullString, SW_NORMAL)
        
    Case 5:
        'show batch log
        txtBatchLog.Visible = True
        
    Case 6:
        
        
    End Select

End Sub

Private Sub lblLinkButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)   'set hand cursor
    lblLinkButton(Index).ForeColor = RGB(255, 0, 0)
    tmrRestoreLinkLabels.Enabled = True
End Sub

Private Sub lblLinkButton_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DoDrop(Data)
    Call cmdNext_Click
End Sub

Private Sub tmrRestoreLinkLabels_Timer()

    Dim nActive As Integer: Dim i As Integer: Dim myX As Integer: Dim myY As Integer

    nActive = lblLinkButton.Count
    For i = lblLinkButton.LBound To lblLinkButton.UBound
        If lblLinkButton(i).ForeColor <> 12582912 Then
            myX = MouseX(lblLinkButton(i).Container.hWnd) * Screen.TwipsPerPixelX
            myY = MouseY(lblLinkButton(i).Container.hWnd) * Screen.TwipsPerPixelY
            If myY < lblLinkButton(i).Top Or myY > lblLinkButton(i).Top + lblLinkButton(i).Height Or myX < lblLinkButton(i).Left Or myX > lblLinkButton(i).Left + lblLinkButton(i).Width Then
                lblLinkButton(i).ForeColor = 12582912
                nActive = nActive - 1
            End If
        Else
            nActive = nActive - 1
        End If
    Next
    If nActive = 0 Then tmrRestoreLinkLabels.Enabled = False

End Sub


Private Sub cmdRemoveSelected_Click()

    Dim n As Integer

    Dim numDeleted As Integer
    numDeleted = 0

    If lwFiles.ListItems.Count = 0 Then Exit Sub

restart:
    For n = 1 To lwFiles.ListItems.Count
        If lwFiles.ListItems(n).Selected = True Then
            lwFiles.ListItems.Remove (n)
            lstFilesComplete.RemoveItem (n - 1)
            numDeleted = numDeleted + 1
            GoTo restart
        End If
    Next

    If numDeleted = 0 Then
        Dim ret As Variant
        ret = MsgBox("No files selected. Do you want to clear the list?", vbQuestion + vbYesNo)
        If ret = vbYes Then
            lwFiles.ListItems.Clear
            lstFilesComplete.Clear
        End If
    End If
    
    lblDocsToProcess = "Documents to process (" & lstFilesComplete.ListCount & "):"

    Call validateFrameInput 'enable/disable cmdNext

End Sub

Private Sub cmdRemoveAll_Click()
    lwFiles.ListItems.Clear
    lstFilesComplete.Clear
    Call validateFrameInput 'enable/disable cmdNext
End Sub

Private Sub chkBatchMode_Click()

    If currentFrame <> FRAME_INPUT And currentFrame <> FRAME_BATCH_INPUT Then Exit Sub

    If chkBatchMode.Value = 1 Then
        currentFrame = FRAME_BATCH_INPUT
    Else
        currentFrame = FRAME_INPUT
    End If

    updateFrames
    
End Sub

Private Sub chkCreateFolders_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DoDrop(Data)
End Sub

Private Sub cmdInput_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DoDrop(Data)
End Sub

Private Sub cmdOutput_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DoDrop(Data)
End Sub


Private Sub cmdAdd_Click()

    ' api open dialog box '''''''''
    Dim tOPENFILENAME As OPENFILENAME
    Dim lResult As Long
    Dim vFiles As Variant
    Dim lIndex As Long, lStart As Long
    Dim sAddFolder As String

    With tOPENFILENAME
        .Flags = OFN_ALLOWMULTISELECT Or OFN_EXPLORER Or OFN_FILEMUSTEXIST Or OFN_LONGNAMES
        .hwndOwner = hWnd
        .nMaxFile = 32768
        .lpstrFilter = "Supported Documents" & Chr(0) & "*.doc;*.dot;*.ppt;*.pot;*.pps;*.docm;*.docx;*.dotm;*.dotx;*.potm;*.potx;*.pptm;*.pptx;*.ppsm;*.ppsx;*.sldm;*.xlsm;*.xlsx;*.xltm;*.xltx;*.odt;*.odp;*.ods;*.ott;*.odg;*.ots;*.otp;*.otg;*.pages;*.numbers;*.template;*.sxw;*.stw;*.sxc;*.stc;*.sxi;*.sti;*.epub;*.cbz;*.fb2;*.xps;*.oxps;*.dwfx;*.chm;*.swf" & Chr(0) & "All Files" & Chr(0) & "*.*" & Chr(0) & Chr(0)
        .lpstrFile = Space(.nMaxFile - 1) & Chr(0)
        .lpstrInitialDir = sLastDir
        .lStructSize = Len(tOPENFILENAME)
    End With

    cmdAdd.Refresh
    cmdRemoveSelected.Refresh
    cmdBatchOutput.Refresh

    lResult = GetOpenFileName(tOPENFILENAME)

    If lResult > 0 Then
        
        Screen.MousePointer = 11
        
        With tOPENFILENAME
            vFiles = Split(Left(.lpstrFile, InStr(.lpstrFile, Chr(0) & Chr(0)) - 1), Chr(0))
        End With

        If UBound(vFiles) = 0 Then
            AddFile (vFiles(0))
            sLastDir = RemoveFile(vFiles(0))
        ElseIf UBound(vFiles) > 0 Then
            sAddFolder = vFiles(0)
            If Right$(sAddFolder, 1) <> "\" Then sAddFolder = sAddFolder & "\"
            sLastDir = sAddFolder
            For lIndex = 1 To UBound(vFiles)
                AddFile (sAddFolder & "\" & vFiles(lIndex))
            Next
        End If
    
        Call validateFrameInput 'enable/disable cmdNext
        Screen.MousePointer = 0
    
    End If

End Sub


Private Sub cmdBatchOutput_Click()

    Dim sFolder As String
    
    If Right(txtBatchOutput, 1) = "\" And Right(txtBatchOutput, 2) <> ":\" Then txtBatchOutput = Left(txtBatchOutput, Len(txtBatchOutput) - 1)   'stip right slash if present
    sFolder = txtBatchOutput
    sFolder = BrowseForFolder(Me, "Select output folder:", , , sFolder, , , , , sysFolders.NetHood)

    cmdAdd.Refresh
    cmdRemoveSelected.Refresh
    cmdBatchOutput.Refresh
    
    If sFolder <> "" Then txtBatchOutput = sFolder

End Sub


Private Function AddFile(toAdd As String) As Boolean

    Dim itmX As ListItem
    Dim tmp As String

    On Error GoTo errHandler

    toAdd = Replace(toAdd, "\\", "\")

    If Dir(toAdd, vbArchive + vbHidden + vbNormal + vbReadOnly + vbSystem) <> "" Then   'file must exist and not be dir

        'check for duplicates
        Dim n As Integer
        For n = 0 To lstFilesComplete.ListCount - 1
            If lstFilesComplete.List(n) = toAdd Then
                AddFile = True
                Exit Function
            End If
        Next

        'add
        tmp = Mid(toAdd, InStrRev(toAdd, "\", -1, vbTextCompare) + 1)
        'If (Right(LCase(tmp), 4) = ".pdf") Then tmp = Left(tmp, Len(tmp) - 4)
        Set itmX = lwFiles.ListItems.Add(, , tmp, , 1)
        lstFilesComplete.AddItem (toAdd)
        
    End If
    
    AddFile = True
    
    lblDocsToProcess = "Documents to process (" & lstFilesComplete.ListCount & "):"
    If lstFilesComplete.ListCount Mod 9 = 0 Then DoEvents
    
    Exit Function


errHandler:
    Debug.Print "Error in: AddFile(" & toAdd & ")"
    Call MsgBox("Error adding file...", vbCritical)
    AddFile = False

End Function

Public Sub RecursiveGetFolderContent(ByVal sFol As String, ByRef OutputList As clsListEmu, bFiles As Boolean, bFolders As Boolean, Optional bRecurse As Boolean = True)

    If bHaveFileSystemObject = False Then Exit Sub
    
    Dim fld, tFld, tFil
    Dim Filename As String
    
    On Error GoTo errHandler
    Set fld = myFso.GetFolder(sFol)
    Filename = Dir(myFso.BuildPath(fld.Path, "*.*"), vbNormal)
    While Len(Filename) <> 0
       If bFiles = True Then OutputList.AddItem myFso.BuildPath(fld.Path, Filename)
       Filename = Dir()  ' Get next file
       'DoEvents
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

Private Sub doExtraction()

    Dim i As Long
    Set m_cUnzip = New cUnzip
    
    On Error GoTo errHandler

    Screen.MousePointer = 11

    'validate input boxes
    If Right(documentOptions(iCurrentDocument).txtOutput, 1) = "\" Then documentOptions(iCurrentDocument).txtOutput = Left(documentOptions(iCurrentDocument).txtOutput, Len(documentOptions(iCurrentDocument).txtOutput) - 1)   'strip right slash if present
    
    DoEvents
    Sleep (1000) 'intensional delay to make sure it does not go too quickly on fast machines...

    'convert office97 to OpenXML ''''
    
    Dim sTmp As String
    sTmp = Right(LCase(documentOptions(iCurrentDocument).txtInput), 4)
    Dim sTmpFile As String
    sTmpFile = ""
    
    
    Dim returnValue As Long
    Dim sCommandLine As String
    Dim sExt
'    Dim iDemoCount As Long
    Dim lstExtractedFiles As New clsListEmu
    documentOptions(iCurrentDocument).iImagesExtracted = 0
    Dim sFileName As String
    Dim n As Long
    
    If sTmp = ".doc" Or sTmp = ".dot" Or sTmp = ".ppt" Or sTmp = ".pps" Or sTmp = ".pot" Then

        If sTmp = ".doc" Then
            sTmpFile = tempFolder & "\officewiz_tmp.docx"
            sCommandLine = Chr(34) & App.Path & "\b2x\doc2x.exe"" """ & documentOptions(iCurrentDocument).txtInput & """ -o """ & sTmpFile & """"
        End If
        If sTmp = ".dot" Then
            sTmpFile = tempFolder & "\officewiz_tmp.dotx"
            sCommandLine = Chr(34) & App.Path & "\b2x\doc2x.exe"" """ & documentOptions(iCurrentDocument).txtInput & """ -o """ & sTmpFile & """"
        End If
        If sTmp = ".ppt" Or sTmp = ".pps" Or sTmp = ".pot" Then
            sTmpFile = tempFolder & "\officewiz_tmp.pptx"
            sCommandLine = Chr(34) & App.Path & "\b2x\ppt2x.exe"" """ & documentOptions(iCurrentDocument).txtInput & """ -o """ & sTmpFile & """"
        End If
        Debug.Print sCommandLine

        ChDrive (App.Path)
        ChDir (App.Path & "\b2x\")
        returnValue = ExecuteCmd(sCommandLine)

        Debug.Print "returvärde: """ & returnValue & """"
        Debug.Print "output: " & sExecuteOutput
        
        documentOptions(iCurrentDocument).txtInput = sTmpFile

    End If
    
    
    'convert fictionbook2 to epub ''''

    If sTmp = ".fb2" Then
        'todo: fileformat: ".fb2.zip"

        sTmpFile = tempFolder & "\officewiz_tmp.epub"
        sCommandLine = Chr(34) & App.Path & "\fb2\fb2toepub.exe"" """ & documentOptions(iCurrentDocument).txtInput & """ """ & sTmpFile & """"
        
        Debug.Print sCommandLine

        ChDrive (App.Path)
        ChDir (App.Path & "\fb2\")
        returnValue = ExecuteCmd(sCommandLine)

        Debug.Print "returvärde: """ & returnValue & """"
        Debug.Print "output: " & sExecuteOutput
        
        documentOptions(iCurrentDocument).txtInput = sTmpFile

    End If
    
    
    ' chm extract '''''''''''''''''''''''''
    
    If sTmp = ".chm" Then
        'hh.exe -decompile t:\out\ t:\file.chm
        'hh.exe is part of windows

        Dim chmTempFolder As String
        chmTempFolder = tempFolder & "\owiztmp"
        If Dir(chmTempFolder, vbDirectory) = "" Then SmartCreateFolder (chmTempFolder)
        
        ChDrive (App.Path)
        ChDir (App.Path & "\fb2\")  'fb2 to make sure captureconsole is found
        returnValue = ExecuteCmd("hh.exe -decompile " & ShortName(chmTempFolder & "\") & " " & ShortName(documentOptions(iCurrentDocument).txtInput))
        
        Set myFso = CreateObject("Scripting.FileSystemObject")
        Call RecursiveGetFolderContent(chmTempFolder, lstExtractedFiles, True, False, True)
        
        For n = 0 To lstExtractedFiles.ListCount - 1
            
            sExt = LCase(Right(lstExtractedFiles.List(n), 4))
            If sExt = ".jpg" Or _
                sExt = "jpeg" Or _
                sExt = ".gif" Or _
                sExt = ".png" Or _
                sExt = ".wmf" Or _
                sExt = ".emf" Or _
                sExt = ".svg" Or _
                sExt = ".bmp" Then

                    Dim sTargetFolder As String
                    sTargetFolder = documentOptions(iCurrentDocument).txtOutput
                    If documentOptions(iCurrentDocument).iCreateFolder = 1 Then
                        sFileName = GetFilename(documentOptions(iCurrentDocument).txtBasename)
                        sTargetFolder = sTargetFolder & "\" & sFileName
                        If Dir(sTargetFolder, vbDirectory) = "" Then SmartCreateFolder (sTargetFolder)
                    End If
                    
                    sFileName = GetFilenameAndExt(lstExtractedFiles.List(n))
                    
                    On Error Resume Next
                    If documentOptions(iCurrentDocument).iCreateFolder = 1 Then
                        Name lstExtractedFiles.List(n) As sTargetFolder & "\" & sFileName
                    Else
                        Name lstExtractedFiles.List(n) As sTargetFolder & "\" & documentOptions(iCurrentDocument).txtBasename & "-" & sFileName
                    End If
                    On Error GoTo errHandler

                    documentOptions(iCurrentDocument).iImagesExtracted = documentOptions(iCurrentDocument).iImagesExtracted + 1
            
            End If
            
        Next
            
        'remove temp files & folder
        myFso.DeleteFolder chmTempFolder, True
        
        GoTo skip_zip

    End If
    '''''''''''''''''''''''''''''''''''''''
    
    ' swf extract '''''''''''''''''''''''''

    If sTmp = ".swf" Then

        'Extract PNGs
        
        ChDrive (App.Path)
        ChDir (App.Path & "\SWFTools\")
        returnValue = ExecuteCmd("swfextract.exe " & ShortName(documentOptions(iCurrentDocument).txtInput))
            'Objects in file D:\032-5.swf:
            ' [-i] 1 Shape: ID(s) 9
            ' [-i] 1 MovieClip: ID(s) 27
            ' [-p] 7 PNGs: ID(s) 2-8
            ' [-s] 18 Sounds: ID(s) 1, 10-26
            ' [-f] 1 Frame: ID(s) 0
            ' [-m] 1 MP3 Soundstream
            
        
        If returnValue = 0 Then
            Debug.Print sExecuteOutput
            
            Dim pos1 As Integer
            Dim pos2 As Integer
            pos1 = InStr(1, sExecuteOutput, "PNG")
            If pos1 = 0 Then GoTo no_pngs
            pos1 = pos1 + 12
            pos2 = InStr(pos1, sExecuteOutput, vbNewLine)
            Dim range As String
            range = Trim(Mid(sExecuteOutput, pos1, pos2 - pos1))
            Dim sNums() As String
            sNums = Split(range, ",")
            Dim iNums() As Integer
            ReDim iNums(0)
            
            For n = LBound(sNums) To UBound(sNums)
                If InStr(sNums(n), "-") > 0 Then
                    
                    Dim nn As Integer
                    Dim r() As String
                    r = Split(sNums(n), "-")
                    For nn = Val(r(0)) To Val(r(1))
                        ReDim Preserve iNums(UBound(iNums) + 1)
                        iNums(UBound(iNums)) = nn
                    Next
                    
                Else
                    ReDim Preserve iNums(UBound(iNums) + 1)
                    iNums(UBound(iNums)) = Val(sNums(n))
                End If
            Next
            
            sFileName = documentOptions(iCurrentDocument).txtOutput
            If documentOptions(iCurrentDocument).iCreateFolder = 1 Then
                sFileName = sFileName & "\" & GetFilename(documentOptions(iCurrentDocument).txtBasename)
            End If
            If UBound(iNums) > 1 Then
                If Dir(sFileName, vbDirectory) = "" Then SmartCreateFolder (sFileName)
            End If
            
            For n = 1 To UBound(iNums)
                If documentOptions(iCurrentDocument).iCreateFolder = 1 Then
                    sCommandLine = "swfextract.exe -p " & iNums(n) & " " & ShortName(documentOptions(iCurrentDocument).txtInput) & " -o """ & sFileName & "\" & iNums(n) & ".png" & """"
                Else
                    sCommandLine = "swfextract.exe -p " & iNums(n) & " " & ShortName(documentOptions(iCurrentDocument).txtInput) & " -o """ & sFileName & "\" & documentOptions(iCurrentDocument).txtBasename & "-" & iNums(n) & ".png" & """"
                End If
                Debug.Print iNums(n)
                Debug.Print vbTab & sCommandLine
                ChDrive (App.Path)
                ChDir (App.Path & "\SWFTools\")
                
                returnValue = ExecuteCmd(sCommandLine)
                If returnValue = 0 Then
                    documentOptions(iCurrentDocument).iImagesExtracted = documentOptions(iCurrentDocument).iImagesExtracted + 1
                End If
                Debug.Print vbTab & returnValue
            
            Next
        
        End If
no_pngs:
        'Extract JPEGs
        '[-j] 130 JPEGs: ID(s) 1, 2, 4, 6, 20, 23, 55, 58, 72, 77, 96, 99, 107, 109, 112, 114, 145, 147, 150, 152, 170, 172, 175, 177, 195, 197, 200, 202, 235, 238, 240, 250, 252, 254, 256, 263, 265, 267, 269, 273, 275, 277, 281, 283, 285, 289, 291, 293, 297, 299, 301, 305, 307, 309, 322, 332, 334, 336, 338, 340, 342, 344, 346, 348, 350, 352, 354, 356, 358, 360, 362, 364, 366, 390, 392, 395, 397, 431, 441, 443, 445, 591, 593, 595, 598, 600, 603, 605, 608, 610, 614, 616, 618, 634, 651, 657, 660, 784, 868, 873, 876, 878, 1157, 1172, 1240, 1242, 1244, 1256, 1315, 1343, 1345, 1374, 1384, 1388, 1390, 1396, 1399, 1400, 1402, 1405, 1408, 1411, 1414, 1417, 1459, 1477, 1480, 1483, 1486, 1489
        
        ChDrive (App.Path)
        ChDir (App.Path & "\SWFTools\")
        returnValue = ExecuteCmd("swfextract.exe " & ShortName(documentOptions(iCurrentDocument).txtInput))
        
        If returnValue = 0 Then
            Debug.Print sExecuteOutput
            
            pos1 = InStr(1, sExecuteOutput, "JPEG")
            If pos1 = 0 Then GoTo skip_zip
            pos1 = pos1 + 12
            pos2 = InStr(pos1, sExecuteOutput, vbNewLine)
            range = Trim(Mid(sExecuteOutput, pos1, pos2 - pos1))
            sNums = Split(range, ",")
            ReDim iNums(0)
            
            For n = LBound(sNums) To UBound(sNums)
                If InStr(sNums(n), "-") > 0 Then
                    
                    r = Split(sNums(n), "-")
                    For nn = Val(r(0)) To Val(r(1))
                        ReDim Preserve iNums(UBound(iNums) + 1)
                        iNums(UBound(iNums)) = nn
                    Next
                    
                Else
                    ReDim Preserve iNums(UBound(iNums) + 1)
                    iNums(UBound(iNums)) = Val(sNums(n))
                End If
            Next
            
            sFileName = documentOptions(iCurrentDocument).txtOutput
            If documentOptions(iCurrentDocument).iCreateFolder = 1 Then
                sFileName = sFileName & "\" & GetFilename(documentOptions(iCurrentDocument).txtBasename)
            End If
            If UBound(iNums) > 1 Then
                If Dir(sFileName, vbDirectory) = "" Then SmartCreateFolder (sFileName)
            End If
            
            For n = 1 To UBound(iNums)
                If documentOptions(iCurrentDocument).iCreateFolder = 1 Then
                    sCommandLine = "swfextract.exe -j " & iNums(n) & " " & ShortName(documentOptions(iCurrentDocument).txtInput) & " -o """ & sFileName & "\" & iNums(n) & ".jpg" & """"
                Else
                    sCommandLine = "swfextract.exe -j " & iNums(n) & " " & ShortName(documentOptions(iCurrentDocument).txtInput) & " -o """ & sFileName & "\" & documentOptions(iCurrentDocument).txtBasename & "-" & iNums(n) & ".jpg" & """"
                End If
                Debug.Print iNums(n)
                Debug.Print vbTab & sCommandLine
                ChDrive (App.Path)
                ChDir (App.Path & "\SWFTools\")
                
                returnValue = ExecuteCmd(sCommandLine)
                If returnValue = 0 Then
                    documentOptions(iCurrentDocument).iImagesExtracted = documentOptions(iCurrentDocument).iImagesExtracted + 1
                End If
                Debug.Print vbTab & returnValue
            
            Next
        
        End If
        
        
        
        
        GoTo skip_zip

    End If
    '''''''''''''''''''''''''''''''''''''''

    'zip extract
    
    ChDrive (App.Path)
    ChDir (App.Path)
    
    m_cUnzip.ZipFile = documentOptions(iCurrentDocument).txtInput
    m_cUnzip.Directory
    For i = 1 To m_cUnzip.FileCount

        m_cUnzip.FileSelected(i) = False
        sExt = LCase(Right(m_cUnzip.Filename(i), 4))
        If sExt = ".jpg" Or _
            sExt = "jpeg" Or _
            sExt = ".gif" Or _
            sExt = ".png" Or _
            sExt = ".tif" Or _
            sExt = "tiff" Or _
            sExt = ".wmf" Or _
            sExt = ".emf" Or _
            sExt = ".svg" Or _
            sExt = ".bmp" Then
                If InStr(LCase(m_cUnzip.FileDirectory(i)), "thumbnail") = 0 And InStr(LCase(m_cUnzip.FileDirectory(i)), "thumbs") = 0 And InStr(LCase(m_cUnzip.FileDirectory(i)), "quicklook") = 0 And InStr(LCase(m_cUnzip.FileDirectory(i)), "previews") = 0 Then
                    m_cUnzip.FileSelected(i) = True
                    lstExtractedFiles.AddItem (m_cUnzip.Filename(i))
                    documentOptions(iCurrentDocument).iImagesExtracted = documentOptions(iCurrentDocument).iImagesExtracted + 1
                End If
        End If
        
    Next i

    If documentOptions(iCurrentDocument).iImagesExtracted > 0 Then

        'extract files
        bFailedOnPasswordFile = False
        m_cUnzip.UnzipFolder = documentOptions(iCurrentDocument).txtOutput  'katalogen skapas automatiskt om den inte finns
        If documentOptions(iCurrentDocument).iCreateFolder = 1 Then
            sFileName = GetFilename(documentOptions(iCurrentDocument).txtBasename)
            m_cUnzip.UnzipFolder = m_cUnzip.UnzipFolder & "\" & sFileName
        End If

        If Right(m_cUnzip.UnzipFolder, 1) <> "\" Then m_cUnzip.UnzipFolder = m_cUnzip.UnzipFolder & "\"

        m_cUnzip.OverwriteExisting = True
        m_cUnzip.Unzip
    
        If bFailedOnPasswordFile = True Then
            documentOptions(iCurrentDocument).iImagesExtracted = 0
        End If
        
    End If


    DoEvents
    

    Sleep (1000) 'intensional delay to make sure it does not go too quickly on fast machines...


    'add basename first in file name
    If documentOptions(iCurrentDocument).iCreateFolder = 0 Then
        For i = 0 To lstExtractedFiles.ListCount - 1
            On Error Resume Next
            Name m_cUnzip.UnzipFolder & "\" & lstExtractedFiles.List(i) As m_cUnzip.UnzipFolder & "\" & documentOptions(iCurrentDocument).txtBasename & "-" & lstExtractedFiles.List(i)
            On Error GoTo errHandler
        Next
    End If
    

    'remove temp file
    If sTmpFile <> "" Then
        On Error Resume Next
        Kill sTmpFile
        On Error GoTo errHandler
    End If


skip_zip:
    Select Case documentOptions(iCurrentDocument).iImagesExtracted
    Case 0:
        documentOptions(iCurrentDocument).sExecuteInfo = "No images to extract..."
        imgDestinationFolder.Visible = False
        lblLinkButton(3).Visible = False
    Case 1:
        documentOptions(iCurrentDocument).sExecuteInfo = "1 image extracted!"
    Case Else
         documentOptions(iCurrentDocument).sExecuteInfo = documentOptions(iCurrentDocument).iImagesExtracted & " images extracted!"
    End Select

    Screen.MousePointer = 0
    Exit Sub

errHandler:
    Screen.MousePointer = 0
    documentOptions(iCurrentDocument).sExecuteInfo = "A program error occurred..."
    documentOptions(iCurrentDocument).bError = True
    imgDestinationFolder.Visible = False
    lblLinkButton(3).Visible = False

End Sub



