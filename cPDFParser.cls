VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPDFParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===================================================================
'             _     ____  ____  _____ ____
'      __   _| |__ |  _ \|  _ \|  ___|  _ \ __ _ _ __ ___  ___ _ __
'      \ \ / / '_ \| |_) | | | | |_  | |_) / _` | '__/ __|/ _ \ '__|
'       \ V /| |_) |  __/| |_| |  _|_|  __/ (_| | |  \__ \  __/ |
'        \_/ |_.__/|_|   |____/|_| / |_/ _ \__,_|_|  |___/\___|_|
'                                  | || | | |
'                                  | || |_| |
'                                  |_(_)___/
'
'===================================================================
' Descrizione.....: Classe per il parsing di un file PDF
'                   Class to parse a PDF file
' Nome dei File...: vbPDFParser.cls
' Data............: 24/06/2008
' Versione........: 1.0
' Sistema.........: Visual Basic 6.0 Pro - SP 6
' Testato su......: Windows XP Professional - SP 2
' Scritto da......: Luigi Micco
' E-Mail..........: l.micco(at)tiscali.it
'===================================================================
'===================================================================
' (C) 2008  - L'uso di questo software è consentito solo su espressa
'             autorizzazione dell'autore. Non puo' essere copiato o
'             ridistribuito, ne' integralmente, ne' parzialmente.
'
'             The use of this software is allowed only on express
'             authorization of the author. It's cannot be copied or
'             redistributed, neither integrally, neither partially.
'===================================================================

Option Explicit

Private mvarImpFileNumber As Integer
Private mvarFilename As String
Private mvarPages() As Long
Private mvarObjOffset() As Long
Private mvarDocVersion As Long
Private mvarPageCount As Integer
Private mvarDocTitle As String
Private mvarDocAuthor As String
Private mvarDocProducer As String
Private mvarDocCreator As String
Private mvarDocSubject As String
Private mvarDocKeywords As String
Private mvarDocDate As String
Private mvarProtection As Long
Private mvarEncrypt As Long

Private CodeBuffer As String
Private CodeCursor As Long

Private iNumIterations  'buggfix

'---------------------------------------------------------------------------------------
' Proprietà : Author
' Risultato : Restituisce l'autore del documento
'---------------------------------------------------------------------------------------
Public Property Get Author() As String
  Author = mvarDocAuthor
End Property


Private Sub Class_Terminate()
  If mvarImpFileNumber <> 0 Then Close #mvarImpFileNumber
End Sub

'---------------------------------------------------------------------------------------
' Proprietà : CreationDate
' Risultato : Restituisce la data di creazione del documento (yyyymmdd)
'---------------------------------------------------------------------------------------
Public Property Get CreationDate() As String
    If Trim(mvarDocDate) <> "" Then
        CreationDate = Mid$(mvarDocDate, 9, 2) & "/" & Mid$(mvarDocDate, 7, 2) & "/" & Mid$(mvarDocDate, 3, 4)
        
        'buggfix för fåtal filer som har ett helt annat format. inte helt perfekt, men undviker att man bara returnerar skräp
        Dim tmp As String
        tmp = Replace(CreationDate, "/", "")
        If Not IsNumeric(tmp) Then
            CreationDate = mvarDocDate
        End If
        
    Else
        CreationDate = ""
    End If
End Property

'---------------------------------------------------------------------------------------
' Proprietà : Creator
' Risultato : Restituisce il creatore del documento
'---------------------------------------------------------------------------------------
Public Property Get Creator() As String
  Creator = mvarDocCreator
End Property

Public Function GetInfo(FileName As String) As Boolean
  
  Dim lngPos As Long
  Dim lngStart As Long
  Dim lngEnd As Long
  Dim strText As String
  Dim lngXREFPos As Long
  Dim lngPageObj As Long
  Dim ObjRoot As Long
  Dim ObjInfo As Long
  Dim blnEncrypted As Boolean
  
  If mvarFilename = FileName Then
    GetInfo = True
    Exit Function
  End If

  mvarImpFileNumber = FreeFile
  mvarFilename = FileName

  Open FileName For Binary Access Read As #mvarImpFileNumber
  
  ' Legge la versione del documento
  strText = Space$(8)
  Get #mvarImpFileNumber, , strText
  mvarDocVersion = Val(Right$(strText, 1))
  
  lngPos = LOF(mvarImpFileNumber)
  strText = Space$(1025)
  Seek #mvarImpFileNumber, lngPos - 1024
  Get #mvarImpFileNumber, , strText
  
  lngStart = InStr(strText, "startxref") + 9
  lngEnd = InStr(lngStart, strText, "%%EOF")
  If lngStart = 9 Or lngEnd = 0 Then    'buggfix
    GetInfo = False
    Close #mvarImpFileNumber
    mvarImpFileNumber = 0
    Exit Function
  End If
  lngXREFPos = Val(Mid$(strText, lngStart, lngEnd - lngStart))
  ReDim mvarObjOffset(0)
  ReDim mvarObjNewId(0)
  
  iNumIterations = 0    'buggfix för att den inte ska hänga sig i nästa
  blnEncrypted = zzPrsReadXREF(lngXREFPos, ObjRoot, ObjInfo)
  
  mvarProtection = 0
  If blnEncrypted Then
    ' Legge le informazione dell'oggetto Encrypt
    strText = zzPrsReadParam(mvarEncrypt)
    zzPrsInitBuffer strText
    'Debug.Print "-------------------"
    'Debug.Print mvarFilename
    'Debug.Print strText
    'Debug.Print "-EOF-"
    If zzPrsSeekToken("V", 1) Then mvarProtection = zzPrsReadNum
  Else
    ' Legge le informazione dell'oggetto INFO
    strText = zzPrsReadParam(ObjInfo)
    zzPrsInitBuffer strText
    If zzPrsSeekToken("Title", 1) Then mvarDocTitle = zzPrsReadValue
    If zzPrsSeekToken("Author", 1) Then mvarDocAuthor = zzPrsReadValue
    If zzPrsSeekToken("Creator", 1) Then mvarDocCreator = zzPrsReadValue
    If zzPrsSeekToken("Producer", 1) Then mvarDocProducer = zzPrsReadValue
    If zzPrsSeekToken("Subject", 1) Then mvarDocSubject = zzPrsReadValue
    If zzPrsSeekToken("CreationDate", 1) Then mvarDocDate = zzPrsReadValue
    If zzPrsSeekToken("Keywords", 1) Then mvarDocKeywords = zzPrsReadValue
    
    ' Legge le informazioni dell'oggetto ROOT
    strText = zzPrsReadParam(ObjRoot)
    zzPrsInitBuffer strText
    If zzPrsSeekToken("Pages", 1) Then
      lngPageObj = zzPrsReadRef
      zzPrsReadPagesTree lngPageObj
    End If
    
  End If
  GetInfo = True

    Close #mvarImpFileNumber
    mvarImpFileNumber = 0

End Function

'---------------------------------------------------------------------------------------
' Sub       : GetObj
' Risultato : Restituisce l'oggetto
'---------------------------------------------------------------------------------------
Public Sub GetObj(ObjId As Long, _
                  ByRef Param As String, _
                  ByRef Stream() As Byte, _
                  ByRef Filter As String)
                  
  Dim blnStream As Boolean

  blnStream = True
  zzPrsReadObj ObjId, Param, blnStream, Stream(), Filter
  If Not blnStream Then ReDim Stream(0)
End Sub

'---------------------------------------------------------------------------------------
' Proprietà : Keywords
' Risultato : Restituisce la stringa delle Keywords del documento
'---------------------------------------------------------------------------------------
Public Property Get Keywords() As String
  Keywords = mvarDocKeywords
End Property

'---------------------------------------------------------------------------------------
' Proprietà : ObjCount
' Risultato : Restituisce il numero di oggetti del documento
'---------------------------------------------------------------------------------------
Public Property Get ObjCount() As Long
  ObjCount = UBound(mvarObjOffset) - 1
End Property

'---------------------------------------------------------------------------------------
' Proprietà : Protection
' Risultato : Restituisce il tipo di protezione associato al documento
'---------------------------------------------------------------------------------------
Public Property Get Protection() As Long
  Protection = mvarProtection
End Property

'---------------------------------------------------------------------------------------
' Proprietà : ObjOffset
' Risultato : Restituisce l'offset di ogni singolo oggetto
'---------------------------------------------------------------------------------------
Public Property Get ObjOffset(ObjIndex As Long) As Long
  ObjOffset = mvarObjOffset(ObjIndex)
End Property

'---------------------------------------------------------------------------------------
' Proprietà : PageCount
' Risultato : Restituisce il numero di pagine
'---------------------------------------------------------------------------------------

Public Property Let PageCount(nPages As Long)
  mvarPageCount = nPages
End Property


Public Property Get PageCount() As Long
  PageCount = mvarPageCount
End Property


'---------------------------------------------------------------------------------------
' Proprietà : PageObj
' Risultato : Restituisce il numero dell'oggetto che contiene la pagina
'---------------------------------------------------------------------------------------
Public Property Get PageObj(PageNumber As Long) As Long
  PageObj = mvarPages(PageNumber)
End Property

'---------------------------------------------------------------------------------------
' Proprietà : Producer
' Risultato : Restituisce il nome dell'applicativo che ha generato il documento
'---------------------------------------------------------------------------------------
Public Property Get Producer() As String
  Producer = mvarDocProducer
End Property

'---------------------------------------------------------------------------------------
' Proprietà : Subject
' Risultato : Restituisce l'oggetto del documento
'---------------------------------------------------------------------------------------
Public Property Get Subject() As String
  Subject = mvarDocSubject
End Property

'---------------------------------------------------------------------------------------
' Proprietà : Title
' Risultato : Restituisce il titolo del documento
'---------------------------------------------------------------------------------------
Public Property Get Title() As String
  Title = mvarDocTitle
End Property

'---------------------------------------------------------------------------------------
' Proprietà : Version
' Risultato : Restituisce la versione del documento
'---------------------------------------------------------------------------------------
Public Property Get Version() As String
  Version = "PDF 1." & mvarDocVersion
End Property

Private Sub zzPrsInitBuffer(ByRef code As String)
  CodeBuffer = code
  CodeCursor = 1
End Sub

Private Function zzPrsReadNum() As Variant
  Dim strTemp As String
  
  zzPrsSkipDummy
  While IsNumeric(Mid(CodeBuffer, CodeCursor, 1)) Or Mid(CodeBuffer, CodeCursor, 1) = "-" Or Mid(CodeBuffer, CodeCursor, 1) = "."
    strTemp = strTemp & Mid(CodeBuffer, CodeCursor, 1)
    CodeCursor = CodeCursor + 1
  Wend
  zzPrsReadNum = Val(strTemp)
End Function

Private Sub zzPrsReadObj(ObjId As Long, ByRef Param As String, _
                         ByRef Stream As Boolean, _
                         ByRef OutBuf() As Byte, ByRef Filter As String)

  Dim strBuffer As String * 1024
  Dim blnFlag As Boolean
  Dim lngStream As Long
  Dim lngLength As Long
  Dim lngOffset As Long
  
  Dim strTemp As String
  Dim lngStartObj As Long
  Dim lngEndObj As Long
  
  Dim tmpCodeBuffer As String
  Dim tmpCodeCursor As Long
  
  tmpCodeBuffer = CodeBuffer
  tmpCodeCursor = CodeCursor
  
  Seek #mvarImpFileNumber, mvarObjOffset(ObjId) + 1
  blnFlag = True
  Do While blnFlag
'DoEvents
    Get #mvarImpFileNumber, , strBuffer
    
    ' Trova l'inizio dell'oggetto
    If lngStartObj = 0 Then
'' 1.1
      lngStartObj = InStr(strBuffer, "obj") + 4
      
      If Asc(Mid$(strBuffer, lngStartObj, 1)) = 10 Then lngStartObj = lngStartObj + 1
      If Mid$(strBuffer, lngStartObj, 2) = vbCrLf Then
        lngStartObj = lngStartObj + 2
      ElseIf Mid$(strBuffer, lngStartObj, 1) = vbCr Then
        lngStartObj = lngStartObj + 1
      End If
      strBuffer = Mid$(strBuffer, lngStartObj)
    End If
    
    ' Trova la fine dell'oggetto o l'inizio dello stream
    lngEndObj = InStr(strBuffer, "endobj" & Chr$(13))
    If lngEndObj = 0 Then lngEndObj = InStr(strBuffer, "endobj" & Chr$(10))
    lngStream = InStr(strBuffer, "stream" & Chr$(10))
    If lngStream = 0 Then lngStream = InStr(strBuffer, "stream" & Chr$(13))
    If (lngEndObj = 0) And (lngStream = 0) Then
      strTemp = strTemp & strBuffer
    Else
      If (lngEndObj = 0) Or ((lngStream <> 0) And (lngStream < lngEndObj)) Then
        lngEndObj = lngStream
      End If
      blnFlag = False
      If lngEndObj > 2 Then 'buggfix
            If Mid$(strBuffer, lngEndObj - 2, 2) = vbCrLf Then
              lngEndObj = lngEndObj - 2
            ElseIf Mid$(strBuffer, lngEndObj - 1, 1) = vbCr Then
              lngEndObj = lngEndObj - 1
            End If
      ElseIf lngEndObj = 2 Then 'buggfix
            If Mid$(strBuffer, lngEndObj - 1, 1) = vbCr Then
              lngEndObj = lngEndObj - 1
            End If
      End If
      strTemp = strTemp & Left$(strBuffer, lngEndObj - 1) ' + 6)
          
    End If
  Loop
  
  Param = strTemp
  
  If Stream Then
    zzPrsInitBuffer strTemp
    If zzPrsSeekToken("Filter", 1) Then Filter = zzPrsReadValue
    If zzPrsSeekToken("Length", 1) Then
      Stream = True
      lngStartObj = CodeCursor
      lngLength = Val(zzPrsReadRecurseRef)
      
      lngEndObj = CodeCursor
      strTemp = Left$(strTemp, lngStartObj) & " " & _
                CStr(lngLength) & " " & _
                Mid$(strTemp, lngEndObj)
      Seek #mvarImpFileNumber, mvarObjOffset(ObjId) + 1
      blnFlag = True
      Do While blnFlag
'DoEvents
        Get #mvarImpFileNumber, , strBuffer
        
        ' Trova l'inizio dello stream
        lngStream = InStr(strBuffer, "stream")
        If (lngStream <> 0) Then
          If Mid$(strBuffer, lngStream + 6, 1) = Chr$(13) Then
            lngOffset = lngStream + 7
          Else
            lngOffset = lngStream + 6
          End If
          blnFlag = False
        
          ReDim OutBuf(lngLength - 1)
          Seek #mvarImpFileNumber, mvarObjOffset(ObjId) + 1 + lngOffset
          Get #mvarImpFileNumber, , OutBuf()
        Else
          lngOffset = lngOffset + 1024
        End If
      Loop
    Else
      Stream = False
    End If
    
  End If
  CodeBuffer = tmpCodeBuffer
  CodeCursor = tmpCodeCursor
  
End Sub

Private Sub zzPrsReadPagesTree(PagesObj As Long)
  Dim strText As String
  Dim strType As String
  Dim Temp() As Long
  Dim i As Long
  
  strText = zzPrsReadParam(PagesObj)
  zzPrsInitBuffer strText
  
    If zzPrsSeekToken("Count", 1) Then
        mvarPageCount = zzPrsReadValue
    End If

  
'  If zzPrsSeekToken("Kids", 1) Then
'    zzPrsReadRefArray temp()
'    For i = 1 To UBound(temp)
'      strText = zzPrsReadParam(temp(i))
'      zzPrsInitBuffer strText
'      If zzPrsSeekToken("Type", 1) Then
'        strType = zzPrsReadValue
'        If strType = "/Page" Then
'          mvarPageCount = mvarPageCount + 1
'          ReDim Preserve mvarPages(mvarPageCount)
'          mvarPages(mvarPageCount) = temp(i)
'        ElseIf strType = "/Pages" Then
'          zzPrsReadPagesTree temp(i)
'        End If
'      End If
'    Next
'  End If

End Sub

Private Function zzPrsReadParam(Index As Long) As String
  Dim Filter As String
  Dim Stream() As Byte
  Dim blnFlag As Boolean
  Dim strTemp As String
  
  blnFlag = False
  Call zzPrsReadObj(Index, strTemp, blnFlag, Stream(), Filter)
  zzPrsReadParam = strTemp
End Function

' Legge il valore a cui punta il riferimento indiretto
Private Function zzPrsReadRecurseRef() As String
  Dim lngTemp As Long
  Dim lngTempR As Long
  Dim strText As String
  
  
  lngTemp = zzPrsReadNum
  zzPrsSkipDummy
  If IsNumeric(Mid(CodeBuffer, CodeCursor, 1)) Then
    lngTempR = zzPrsReadNum
    zzPrsSkipDummy
    If Mid(CodeBuffer, CodeCursor, 1) = "R" Then strText = zzPrsReadParam(lngTemp)
  Else
    strText = CStr(lngTemp)
  End If
  
  strText = Replace(Replace(strText, vbCr, " "), vbLf, " ")
  Do While InStr(strText, "  ") <> 0
'DoEvents
    strText = Replace(strText, "  ", " ")
  Loop
  zzPrsReadRecurseRef = strText
  
End Function

' Legge l'ID dell'oggetto a cui si riferisce il riferimento indiretto
Private Function zzPrsReadRef() As Long
  Dim lngTemp As Long
  Dim lngTempR As Long
  
  lngTemp = zzPrsReadNum
  zzPrsSkipDummy
  lngTempR = zzPrsReadNum
  zzPrsSkipDummy
  If Mid(CodeBuffer, CodeCursor, 1) = "R" Then
    zzPrsReadRef = lngTemp
    CodeCursor = CodeCursor + 1
  End If
End Function

' Legge un array di riferimenti indiretti
Private Sub zzPrsReadRefArray(ByRef aObj() As Long)
  Dim lngObjNum As Long
  Dim i As Long
  Dim lngEnd As Long
  
  i = 0
  zzPrsSkipDummy
  If Mid(CodeBuffer, CodeCursor, 1) = "[" Then
    CodeCursor = CodeCursor + 1
    lngEnd = InStr(CodeCursor, CodeBuffer, "]")
    While CodeCursor <= (lngEnd - 5)
'DoEvents
      lngObjNum = zzPrsReadRef
      If lngObjNum <> 0 Then
        i = i + 1
        ReDim Preserve aObj(i)
        aObj(i) = lngObjNum
      End If
    Wend
  Else
    lngObjNum = zzPrsReadRef
    If lngObjNum <> 0 Then
      ReDim aObj(1)
      aObj(1) = lngObjNum
    End If
  End If

End Sub

' Legge il valore presente alla posizione del cursore
Private Function zzPrsReadValue(Optional Recursive As Boolean = False) As String
  Dim strTemp As String
  Dim strValue As String
  Dim lngStart As Long
  Dim lngEnd As Long
  Dim i As Long
  Dim strChar As String
  Dim lngTemp As Long
  Dim strText As String
  Dim lngCount As Long
  Dim lngTempR As Long
    
  zzPrsSkipDummy
  If Mid$(CodeBuffer, CodeCursor, 1) = "/" Then
  
    lngStart = CodeCursor
    CodeCursor = CodeCursor + 1
    Do While (Mid$(CodeBuffer, CodeCursor, 1) <> "[") And _
       (Mid$(CodeBuffer, CodeCursor, 1) <> " ") And _
       (Mid$(CodeBuffer, CodeCursor, 1) <> "/") And _
       (Mid$(CodeBuffer, CodeCursor, 1) <> "(") And _
       (Mid$(CodeBuffer, CodeCursor, 1) <> "<") And _
       (Mid$(CodeBuffer, CodeCursor, 1) <> ">") And _
       (Mid$(CodeBuffer, CodeCursor, 1) <> vbCr) And _
       (Mid$(CodeBuffer, CodeCursor, 1) <> vbLf)
      CodeCursor = CodeCursor + 1
'DoEvents
    Loop
    zzPrsReadValue = Mid$(CodeBuffer, lngStart, CodeCursor - lngStart)
  
  ElseIf Mid$(CodeBuffer, CodeCursor, 1) = "[" Then
    ' Legge un array di valori
    zzPrsSkipDummy
    If Mid(CodeBuffer, CodeCursor, 1) = "[" Then
      lngStart = CodeCursor
      lngCount = 1
      CodeCursor = CodeCursor + 1
      While lngCount > 0
'DoEvents
        If Mid(CodeBuffer, CodeCursor, 1) = "[" Then lngCount = lngCount + 1
        If Mid(CodeBuffer, CodeCursor, 1) = "]" Then lngCount = lngCount - 1
        CodeCursor = CodeCursor + 1
      Wend
      strText = Mid$(CodeBuffer, lngStart, CodeCursor - lngStart)
      strText = Replace(Replace(strText, vbCr, " "), vbLf, " ")
      Do While InStr(strText, "  ") <> 0
'DoEvents
        strText = Replace(strText, "  ", " ")
      Loop
    End If
    zzPrsReadValue = strText
  
  ElseIf Mid$(CodeBuffer, CodeCursor, 2) = "<<" Then
    ' Legge un dizionario

    
    zzPrsSkipDummy
    If IsNumeric(Mid(CodeBuffer, CodeCursor, 1)) Then
      lngTemp = zzPrsReadNum
      zzPrsSkipDummy
      lngTempR = zzPrsReadNum
      zzPrsSkipDummy
      If Mid(CodeBuffer, CodeCursor, 1) = "R" Then
        strText = zzPrsReadParam(lngTemp)
        lngStart = InStr(strText, "obj") + 3
        lngEnd = InStr(strText, "endobj") - 1
        strText = Mid$(strText, lngStart, lngEnd - lngStart)
      End If
    ElseIf Mid(CodeBuffer, CodeCursor, 2) = "<<" Then
      lngStart = CodeCursor
      lngCount = 1
      CodeCursor = CodeCursor + 2
      While lngCount > 0
'DoEvents
        If Mid(CodeBuffer, CodeCursor, 2) = "<<" Then lngCount = lngCount + 1
        If Mid(CodeBuffer, CodeCursor, 2) = ">>" Then lngCount = lngCount - 1
        CodeCursor = CodeCursor + 1
      Wend
      strText = Mid$(CodeBuffer, lngStart, (CodeCursor + 1) - lngStart + 1)
    End If
    strText = Replace(Replace(strText, vbCr, " "), vbLf, " ")
    
    Do While InStr(strText, "  ") <> 0
'DoEvents
      strText = Replace(strText, "  ", " ")
    Loop
    zzPrsReadValue = strText
  
  ElseIf (Mid$(CodeBuffer, CodeCursor, 1) = "(") Or (Mid$(CodeBuffer, CodeCursor, 1) = "<") Then
  
    If Mid(CodeBuffer, CodeCursor, 1) = "(" Then
      CodeCursor = CodeCursor + 1
      lngStart = CodeCursor
      While lngEnd = 0
'DoEvents
        If Mid(CodeBuffer, CodeCursor, 2) = "\\" Then CodeCursor = CodeCursor + 2
        If Mid(CodeBuffer, CodeCursor, 2) = "\(" Then CodeCursor = CodeCursor + 2
        If Mid(CodeBuffer, CodeCursor, 2) = "\)" Then CodeCursor = CodeCursor + 2
        If Mid(CodeBuffer, CodeCursor, 1) = ")" Then lngEnd = CodeCursor - 1
        CodeCursor = CodeCursor + 1
      Wend
      strTemp = Mid$(CodeBuffer, lngStart, lngEnd - lngStart + 1)
      zzPrsReadValue = Replace(Replace(Replace(strTemp, "\\", "\"), "\(", "("), "\)", ")")
    ElseIf Mid(CodeBuffer, CodeCursor, 1) = "<" Then
      CodeCursor = CodeCursor + 1
      lngStart = CodeCursor
      CodeCursor = InStr(lngStart, CodeBuffer, ">")
      strTemp = Mid$(CodeBuffer, lngStart, CodeCursor - lngStart)
      CodeCursor = CodeCursor + 1
      For i = 1 To Len(strTemp) - 1 Step 2
        strChar = Chr$(Val("&H" & Mid$(strTemp, i, 2)))
        If (strChar >= " ") And (strChar <= "~") Then strValue = strValue & strChar
      Next
      zzPrsReadValue = strValue
    End If
  
  ElseIf IsNumeric(Mid$(CodeBuffer, CodeCursor, 1)) Then
    If Recursive Then
      zzPrsReadValue = zzPrsReadRecurseRef
    Else
      ' Legge un numero
      lngTemp = zzPrsReadNum
      zzPrsSkipDummy
      If IsNumeric(Mid(CodeBuffer, CodeCursor, 1)) Then
        lngTempR = zzPrsReadNum
        zzPrsSkipDummy
        If Mid(CodeBuffer, CodeCursor, 1) = "R" Then
          zzPrsReadValue = CStr(lngTemp) & " " & CStr(lngTempR) & " R"
        End If
      Else
        zzPrsReadValue = CStr(lngTemp)
      End If
    
    End If
    CodeCursor = CodeCursor + 1
  End If
  
End Function

' Recupera le informazioni dalla tabella 'xref'
Private Function zzPrsReadXREF(XREFStart As Long, _
                               ByRef ObjRoot As Long, ByRef ObjInfo As Long) As Boolean
  Dim strTemp As String
  Dim bytTemp As Byte
  Dim i As Long
  Dim strDummy As String * 20
  Dim lngStart As Long
  Dim lngCount As Long
  Dim lngSize As Long

  Dim blnFlag As Boolean
  
  iNumIterations = iNumIterations + 1
  If iNumIterations > 16 Then
    Debug.Print "iNumIterations > 16 !!!"   'om detta händer är filen kanske korrupt???
    zzPrsReadXREF = False
    Exit Function
  End If
  
  
  Seek #mvarImpFileNumber, (XREFStart + 1)
  
  For i = 1 To 4
    Get #mvarImpFileNumber, , bytTemp
    strTemp = strTemp & Chr$(bytTemp)
  Next
  
  If strTemp = "xref" Then
    
    Do
      Get #mvarImpFileNumber, , bytTemp
    Loop While (bytTemp = 13) Or (bytTemp = 10) Or (bytTemp = 32)
    
    blnFlag = True
    Do While blnFlag
'DoEvents
      strTemp = ""
      Do While IsNumeric(Chr$(bytTemp))
'DoEvents
        strTemp = strTemp & Chr$(bytTemp)
        Get #mvarImpFileNumber, , bytTemp
      Loop
      lngStart = Val(strTemp)
      
      Do While (bytTemp = 32)
'DoEvents
        Get #mvarImpFileNumber, , bytTemp
      Loop
      
      strTemp = ""
      Do While IsNumeric(Chr$(bytTemp))
'DoEvents
        strTemp = strTemp & Chr$(bytTemp)
        Get #mvarImpFileNumber, , bytTemp
      Loop
      lngCount = Val(strTemp)
      
      If (lngStart + lngCount - 1) > UBound(mvarObjOffset()) Then
        ReDim Preserve mvarObjOffset(lngStart + lngCount - 1)
      End If
      
      Do
        Get #mvarImpFileNumber, , bytTemp
      Loop While (bytTemp = 13) Or (bytTemp = 10) Or (bytTemp = 32)
      Seek #mvarImpFileNumber, (Seek(mvarImpFileNumber) - 1)
      
      For i = lngStart To (lngStart + lngCount - 1)
        Get #mvarImpFileNumber, , strDummy
        If mvarObjOffset(i) = 0 Then mvarObjOffset(i) = Val(Left$(strDummy, 10))
      Next
      
      Do
        Get #mvarImpFileNumber, , bytTemp
      Loop While (bytTemp = 13) Or (bytTemp = 10) Or (bytTemp = 32)

      blnFlag = False
      If IsNumeric(Chr$(bytTemp)) Then
        blnFlag = True
      ElseIf Chr$(bytTemp) = "t" Then
        strTemp = Space$(1024)
        Get #mvarImpFileNumber, , strTemp
      
        ' Read the trailer object
        zzPrsInitBuffer "t" & strTemp
        
        If zzPrsSeekToken("Encrypt", 1) Then
          mvarEncrypt = zzPrsReadRef
          zzPrsReadXREF = True
          Exit Function
        End If
                
        
        If zzPrsSeekToken("Size", 1) Then
          lngSize = zzPrsReadNum
          If lngSize > UBound(mvarObjOffset) Then
            ReDim Preserve mvarObjOffset(lngSize)
          End If
        End If

        If (ObjInfo = 0) Then If zzPrsSeekToken("Info", 1) Then ObjInfo = zzPrsReadRef
        If (ObjRoot = 0) Then If zzPrsSeekToken("Root", 1) Then ObjRoot = zzPrsReadRef
        If zzPrsSeekToken("Prev", 1) Then zzPrsReadXREF = zzPrsReadXREF(zzPrsReadNum, ObjRoot, ObjInfo)
      End If
    Loop
  End If

End Function


' Posiziona il cursore all'inizio del token
Private Function zzPrsSeekToken(ByVal Token As String, _
                                 Optional StartPos As Long = 0) As Boolean
  Dim lngTemp As Long

  If StartPos = 0 Then StartPos = CodeCursor
  lngTemp = InStr(StartPos, CodeBuffer, "/" & Token)
  If lngTemp <> 0 Then
    CodeCursor = lngTemp + Len(Token) + 1
    zzPrsSeekToken = True
  End If
  
End Function

' Salta i caratteri non significativi
Private Sub zzPrsSkipDummy()
  While Mid$(CodeBuffer, CodeCursor, 1) = " " Or _
        Mid$(CodeBuffer, CodeCursor, 1) = vbCr Or _
        Mid$(CodeBuffer, CodeCursor, 1) = vbLf Or _
        Mid$(CodeBuffer, CodeCursor, 1) = vbTab
    CodeCursor = CodeCursor + 1
  Wend
End Sub


