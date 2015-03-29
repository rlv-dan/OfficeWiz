Attribute VB_Name = "modShellSort"
'
' Shell Sort
'
' Original code: http://msdn.microsoft.com/library/default.asp?url=/library/en-us/dnoffpro01/html/ABetterShellSortPartI.asp
' Modified by Dan Saeden
'
' Takes an array and sorts the content.
' Handles arrays of number/string/date data types.
' ShellSortStr: Optimized version that only accepts string arrays and is much faster.
' Takes into account lower/upper case and locale settings. (Most other sortings routines will sort a after B, because a comes after B in the ascii table).
' Optional: Can sort Ascending/Descending by adding True/False at end of procedure call.
' Quite fast, although other methods can be faster. Not slow in any way though.
'
' Usage: ShellSort myArray()            'sort myArray()
'        ShellSort myArray() , True     'sort Descending
'        ShellSortStr myStrArray()      'optimized for string arrays
'

Sub ShellSort(vArray As Variant, Optional Decending As Boolean)
    Dim TempVal As Variant
    Dim i As Long, GapSize As Long, CurPos As Long
    Dim FirstRow As Long, LastRow As Long, NumRows As Long
    FirstRow = LBound(vArray)
    LastRow = UBound(vArray)
    NumRows = LastRow - FirstRow + 1
    Do
      GapSize = GapSize * 3 + 1
    Loop Until GapSize > NumRows
    Do
      GapSize = GapSize \ 3
      For i = (GapSize + FirstRow) To LastRow
        CurPos = i
        TempVal = vArray(i)
        Do While CompareResult(vArray(CurPos - GapSize), TempVal, Decending)
          vArray(CurPos) = vArray(CurPos - GapSize)
          CurPos = CurPos - GapSize
          If (CurPos - GapSize) < FirstRow Then Exit Do
        Loop
        vArray(CurPos) = TempVal
      Next
    Loop Until GapSize = 1
End Sub

Private Function CompareResult(Value1 As Variant, Value2 As Variant, Optional Descending As Boolean)
    If IsDate(Value1) And IsDate(Value2) Then
        CompareResult = (CDate(Value1) > CDate(Value2))
    ElseIf IsNumeric(Value1) And IsNumeric(Value2) Then
        CompareResult = (Value1 > Value2)
    Else
        CompareResult = (StrComp(Value1, Value2, vbTextCompare) = 1)
    End If
    CompareResult = CompareResult Xor Descending
End Function


Sub ShellSortStr(vArray() As String)
    Dim TempVal As String
    Dim i As Long, GapSize As Long, CurPos As Long
    Dim FirstRow As Long, LastRow As Long, NumRows As Long
    FirstRow = LBound(vArray)
    LastRow = UBound(vArray)
    NumRows = LastRow - FirstRow + 1
    Do
      GapSize = GapSize * 3 + 1
    Loop Until GapSize > NumRows
    Do
      GapSize = GapSize \ 3
      For i = (GapSize + FirstRow) To LastRow
        CurPos = i
        TempVal = vArray(i)
        Do While (StrComp(vArray(CurPos - GapSize), TempVal, vbTextCompare) = 1)
          vArray(CurPos) = vArray(CurPos - GapSize)
          CurPos = CurPos - GapSize
          If (CurPos - GapSize) < FirstRow Then Exit Do
        Loop
        vArray(CurPos) = TempVal
      Next
    Loop Until GapSize = 1
End Sub

