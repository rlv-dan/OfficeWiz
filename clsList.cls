VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsListEmu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' VB6 List Control emulation class
' By RL Vision

' (!) There is no error checking if you read/write outside array bounds

Private Type listType
    txtText As String
    bSelected As Boolean
End Type

Private clsList() As listType
Private dimmedUpperBound As Long
Private currentUpperBound As Long
'Private bSorted As Boolean

Public Property Get Selected(lngItemNumber As Long) As Boolean
    Selected = clsList(lngItemNumber + 1).bSelected
End Property

Public Property Let Selected(lngItemNumber As Long, bSelected As Boolean)
    clsList(lngItemNumber + 1).bSelected = bSelected
End Property

Public Property Get List(lngItemNumber As Long) As String
    List = clsList(lngItemNumber + 1).txtText
End Property

Public Property Let List(lngItemNumber As Long, sText As String)
    clsList(lngItemNumber + 1).txtText = sText
End Property

Public Function Find(sFind As String) As Long
    'returns the index of the first found item in the list
    Dim n As Long
    Find = -1
    For n = 1 To currentUpperBound
        If clsList(n).txtText = sFind Then
            Find = n - 1
            Exit Function
        End If
    Next
End Function

Public Function ListCount() As Long
    ListCount = currentUpperBound
End Function

Public Sub RemoveItem(lngItemNumber As Long)
    Dim n As Long
    For n = (lngItemNumber + 1) To (currentUpperBound - 1)
        clsList(n).txtText = clsList(n + 1).txtText
        clsList(n).bSelected = clsList(n + 1).bSelected
    Next
    currentUpperBound = currentUpperBound - 1
End Sub

Public Sub AddItem(txtItem As String)
    currentUpperBound = currentUpperBound + 1
    If currentUpperBound > dimmedUpperBound Then
        dimmedUpperBound = dimmedUpperBound + 100
        ReDim Preserve clsList(1 To dimmedUpperBound)
    End If
    clsList(currentUpperBound).txtText = txtItem
End Sub

Public Sub Clear()
    currentUpperBound = 0
    dimmedUpperBound = 0
    Erase clsList()
End Sub

Private Sub Class_Initialize()
    currentUpperBound = 0
    dimmedUpperBound = 0
End Sub


Public Sub ShellSort()

    'optimized version of modShellSort.bas

    If currentUpperBound <= 1 Then Exit Sub

    Dim TempVal As String
    Dim i As Long, GapSize As Long, CurPos As Long
    Dim FirstRow As Long, LastRow As Long, NumRows As Long
    FirstRow = 1
    LastRow = currentUpperBound
    NumRows = LastRow - FirstRow + 1
    Do
      GapSize = GapSize * 3 + 1
'DoEvents
    Loop Until GapSize > NumRows
    Do
      GapSize = GapSize \ 3
      For i = (GapSize + FirstRow) To LastRow
        CurPos = i
        TempVal = clsList(i).txtText
        Do While (StrComp(clsList(CurPos - GapSize).txtText, TempVal, vbTextCompare) = 1)
          clsList(CurPos).txtText = clsList(CurPos - GapSize).txtText
          CurPos = CurPos - GapSize
          If (CurPos - GapSize) < FirstRow Then Exit Do
'DoEvents
        Loop
        clsList(CurPos).txtText = TempVal
      Next
'DoEvents
    Loop Until GapSize = 1


End Sub

Public Sub ShellSort_Descending()

    'optimized version of modShellSort.bas

    If currentUpperBound <= 1 Then Exit Sub

    Dim TempVal As String
    Dim i As Long, GapSize As Long, CurPos As Long
    Dim FirstRow As Long, LastRow As Long, NumRows As Long
    FirstRow = 1
    LastRow = currentUpperBound
    NumRows = LastRow - FirstRow + 1
    Do
      GapSize = GapSize * 3 + 1
    Loop Until GapSize > NumRows
    Do
      GapSize = GapSize \ 3
      For i = (GapSize + FirstRow) To LastRow
        CurPos = i
        TempVal = clsList(i).txtText
        Do While ((StrComp(clsList(CurPos - GapSize).txtText, TempVal, vbTextCompare) = 1) Xor True)
          clsList(CurPos).txtText = clsList(CurPos - GapSize).txtText
          CurPos = CurPos - GapSize
          If (CurPos - GapSize) < FirstRow Then Exit Do
        Loop
        clsList(CurPos).txtText = TempVal
      Next
    Loop Until GapSize = 1


End Sub

