VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SudokuGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit '593 Zeilen
Private mTxtBoxes As Variant '() 'As TextBox
Public mBlockCol As Collection 'ein  Block  von z.B 3*3 Zellen 'mn * mn
Public mLineCol As Collection  'eine Zeile  von z.B 3*3 Zellen
Public mColmCol As Collection  'eine Spalte von z.B 3*3 Zellen
Public ValueCol As Collection  'speichert alle 81 Zellen
Private mGLevel As String 'der Schwierigkeitsgrad, eins, zwei, drei, vier
Private msCheckErr As String   'der ErrorString von Check
'Public UndoRedo As New SGUndoRedo
'Public UndoList As New Collection 'speichert Undoinstanzen von ValueCol (des ganzen Games)
'Public RedoList As New Collection 'speichert RedoInstanzen des ganzen Games
Private mn As Long '2(4*4), 3(9*9), 4(16*16) = 256, also gleich Long!
Public OptionSolve As Long
Private mLastNumVal As Long
Public Event SaveOnSolve(Obj As SudokuGame)
Public GameName As String

Private Sub Class_Initialize()
  OptionSolve = 3
  Set mBlockCol = New Collection
  Set mLineCol = New Collection
  Set mColmCol = New Collection
End Sub

Public Sub NewC(GameWidth As Long)
'we make it a bit fehlertolerant
  'W darf sein:
  '1 dann ist es 2*2*2*2 = 16
  '2 dann ist es 2*2*2*2
  '3 dann ist es 3*3*3*3 =
  '4 dann ist es 4*4*4*4
  '5..15 dann ist es 2*2*2*2
  '16 dann ist es 4*4*4*4
  '17..81 dann ist es 3*3*3*3
  '82..255 dann ist es 4*4*4*4
  Select Case GameWidth
  Case 1, 2, 5 To 8, 10 To 15: mn = 2 ' 4 = 2*2
  Case 3, 9, 17 To 81: mn = 3 ' 9 = 3*3
  Case 4, 82 To 256: mn = 4 ' 16 = 4*4
  Case Else: MsgBox "Mehr Felder als 256 werden in dieser Version nicht unterstützt."
  End Select
  Call InitBLCs
  Set ValueCol = New_VCV2BLC 'New Collection
End Sub
Private Sub InitBLCs() 'bleibt eigentlich immer gleich
Dim BLC As SudokuBLC
Dim i As Long
  For i = 1 To mn * mn '9
    Set BLC = New SudokuBLC
    BLC.mn = mn: BLC.BLCTyp = stBlock
    BLC.Nr = i: mBlockCol.Add BLC
    Set BLC = New SudokuBLC
    BLC.mn = mn: BLC.BLCTyp = stLine
    BLC.Nr = i: mLineCol.Add BLC
    Set BLC = New SudokuBLC
    BLC.mn = mn: BLC.BLCTyp = stColum
    BLC.Nr = i: mColmCol.Add BLC
  Next
End Sub
Public Function New_VCV2BLC() As Collection
'erzeugt eine Collection...
  Set New_VCV2BLC = New Collection
  Call InitNewSuVals(New_VCV2BLC)
  Call InitVal2BLCs(New_VCV2BLC)
End Function

'bei undo-schritt, neue ValCol anlegen
Private Sub InitNewSuVals(ValCol As Collection)
Dim V As SudokuVal
Dim i As Long, n4 As Long
  n4 = mn * mn * mn * mn
  For i = 1 To n4 '81
    Set V = New SudokuVal
    V.Nr = i
    ValCol.Add V
  Next
End Sub

'verknüpft die Val-Zellen mit zugehörigem Block, Spalte, Zeile
Private Sub InitVal2BLCs(ValCol As Collection)
Dim Idx As Long, Offset As Long, jj As Long, ii As Long
Dim h As Long, i As Long, j As Long, k As Long, n2 As Long, n3 As Long
Dim B As SudokuBLC
Dim L As SudokuBLC
Dim C As SudokuBLC
Dim V As SudokuVal
  n2 = mn * mn
  n3 = n2 * mn
  For k = 1 To mn '3
    For j = 1 To mn '3
      jj = j + (k - 1) * mn '3 'für Zeilen und Spalten
      Set B = mBlockCol(jj)
      Call B.Clear
      Offset = ((k - 1) * n3) + ((j - 1) * mn) '3) 'für die Blöcke
      Set L = mLineCol(jj)
      Call L.Clear
      Set C = mColmCol(jj)
      Call C.Clear
      For i = 1 To mn '3
        For h = 1 To mn '3
          ii = h + (i - 1) * mn '3 'für Zeilen und Spalten
      'in die Zeilen speichern
          Idx = ii + (jj - 1) * n2 '9)
          Set V = ValCol(Idx)
          V.LineNr = jj
          L.ValCol.Add V
      'in die Spalten speichern
          Idx = jj + (ii - 1) * n2 '9)
          Set V = ValCol(Idx)
          V.ColmNr = jj
          C.ValCol.Add V
      'in die Blöcke
          Idx = (h + (i - 1) * n2) + Offset
          Set V = ValCol(Idx)
          V.BlckNr = jj
          B.ValCol.Add V
        Next
      Next
    Next
  Next
End Sub

'Public Sub TestLines()
'Dim L As SudokuBLC
'Dim i As Long
'  Clear
'  For Each L In mLineCol
'    i = i + 1
'    Call L.SetVal(i)
'  Next
'  Show
'End Sub
'Public Sub TestColms()
'Dim C As SudokuBLC
'Dim i As Long
'  Clear
'  For Each C In mColmCol
'    i = i + 1
'    Call C.SetVal(i)
'  Next
'  Show
'End Sub
'Public Sub TestBlocks()
'Dim B As SudokuBLC
'Dim i As Long
'  Clear
'  For Each B In mBlockCol
'    i = i + 1
'    Call B.SetVal(i)
'  Next
'  Show
'End Sub
Public Property Get GameLevel() As Long
  If mGLevel = "eins" Then
    GameLevel = 1
  ElseIf mGLevel = "zwei" Then
    GameLevel = 2
  ElseIf mGLevel = "drei" Then
    GameLevel = 3
  ElseIf mGLevel = "vier" Then
    GameLevel = 4
  Else
    GameLevel = 0
  End If
End Property
Public Property Let GameLevel(LngVal As Long)
  If LngVal = 1 Then
    mGLevel = "eins"
  ElseIf LngVal = 2 Then
    mGLevel = "zwei"
  ElseIf LngVal = 3 Then
    mGLevel = "drei"
  ElseIf LngVal = 4 Then
    mGLevel = "vier"
  Else
    mGLevel = "null"
  End If
End Property

'Public Property Get TxtBoxes() As Variant
'  Set TxtBoxes() = mTxtBoxes()
'End Property
Public Sub SetTxtBoxes(TxtB) '() As Variant)
  Set mTxtBoxes = TxtB
  InitVal2TxtB
End Sub
Public Sub InitVal2TxtB()
Dim V As SudokuVal
Dim i As Long ', n4 As Long
  'If Not mTxtBoxes Is Nothing Then
    'n4 = mn * mn * mn * mn
    For i = 1 To ValueCol.Count 'n4 '81
      Set V = ValueCol(i)
      Set V.TxtBox = mTxtBoxes(i)
      Call V.UpdateToolTip
    Next
  'End If
End Sub
'zeigt das gesamte Spiel in den Textboxen an
Public Sub Show()
Dim Val As SudokuVal
Dim BLC As SudokuBLC
  For Each Val In ValueCol
    Val.ToTextBox
    Call Val.UpdateToolTip
  Next
  For Each BLC In mBlockCol
    Call BLC.UpdateToolTip
  Next
  For Each BLC In mLineCol
    Call BLC.UpdateToolTip
  Next
  For Each BLC In mColmCol
    Call BLC.UpdateToolTip
  Next
End Sub

Public Sub Clear()
Dim V As SudokuVal
Dim BLC As SudokuMissPoss
  For Each V In ValueCol
    V.Clear
  Next
  For Each BLC In mBlockCol
    BLC.ClearMissPossCol
  Next
  For Each BLC In mLineCol
    BLC.ClearMissPossCol
  Next
  For Each BLC In mColmCol
    BLC.ClearMissPossCol
  Next
End Sub

Public Function GetAmountOfUnZeroVals() As Long
Dim V As SudokuVal
  For Each V In ValueCol
    If V.Value > 0 Then GetAmountOfUnZeroVals = GetAmountOfUnZeroVals + 1
  Next
End Function
Public Sub Solve()
Dim B As SudokuBLC
Dim L As SudokuBLC
Dim C As SudokuBLC
Dim V As SudokuVal
Dim n4 As Long
Dim CurNumVal As Long
Dim mess As String
Dim RetVal As Long 'VbMsgBoxResult
Dim GTFOV As String
Dim AppHelp As String
  n4 = mn * mn * mn * mn
'nur was machen wenn noch nicht alle Felder ausgefüllt
  CurNumVal = GetAmountOfUnZeroVals
  If CurNumVal = 0 Then
    mess = "No calculation, all cells are empty."
    MsgBox mess
  Else
'nur was machen wenn noch nicht alle Felder ausgefüllt
    If CurNumVal < n4 Then
      If CurNumVal = mLastNumVal Then
        If OptionSolve > 0 Then
          Call GetMissingValInBLC
          Call GetPossibleValsToVals
        End If
        GTFOV = GetTryFillOutVal
        If Len(GTFOV) > 0 Then
          mess = "There was no change since the last solve, try to fill out one cell of the following:" & vbCrLf
          mess = mess & GTFOV
          AppHelp = App.Path & "\" & App.HelpFile '
          RetVal = MsgBox(mess, vbOKOnly Or vbMsgBoxHelpButton, , AppHelp, 0&)
        Else
          If Not CheckMe Then MsgBox msCheckErr
        End If
        'If RetVal = vbOK Then frmAbout.Show 0
        
        mLastNumVal = 0
      'ElseIf CurNumVal = n4 Then
      '  DoTheCheck
        'If Not CheckMe Then MsgBox msCheckErr

      Else
        mLastNumVal = CurNumVal
      
  'natürlich soll nur dann etwas in der Undoliste gespeichert werden,
  'wenn auch tatsächlich eine Aktion durchgeführt wurde.
  '1. die Berechnung erfolgt nur dann, wenn
  '   nicht alle Werte ungleich null sind.
  '2. es kann auch vorkommen, daß die Berechnung läuft,
  '   aber sich überhaupt kein Wert geändert hat.
  '   wann soll man also das Undo vornehmen?
  '   man könnte das Undo wieder löschen, wenn die Berechnung nichts ergab
  '   man sollte mehr mit Messages arbeiten
    'Call SaveCompleteUndo
        RaiseEvent SaveOnSolve(Me)
        DoEvents
'1. Schritt
'in jedem Block, Zeile und Spalte die in diesem
'Block Zeile oder Spalte fehlenden Werte feststellen,
  'alle Blöcke durchwandern
  '1. Schritt immer durchführen
        If OptionSolve > 0 Then
          Call GetMissingValInBLC
          Call GetPossibleValsToVals
        End If
'2. Schritt
'in jeder Zelle die möglichen Werte feststellen, nur diejenigen fehlenden
'Werte sind mögliche Werte, die in allen drei Listen (Block, Zeile, Spalte)
'gleichzeitig vorkommen.
'Ist daraufhin in einer Zelle nur ein Wert möglich,
'dann ist das der gesuchte Wert dieser Zelle
'Bsp:
'Zelle1=0: PosVal: 2,3,4;
'Zelle2=0: PosVal: 3;
'Zelle3=0: PosVal: 2,3,5;
'=> der Gesuchte Wert in Zelle2 ist 3
        If OptionSolve > 1 Then Call GetSinglePosValToCell
    
'3. Schritt
'kommt in der Liste der möglichen Werte, über alle Zellen eines Blockes
'einer Zeile oder einer Spalte, ein Wert überhaupt nur einmal vor,
'dann kann dieser Wert nur in dieser Zelle stehen.
'Bsp.:
'Zelle1=0: PosVal: 2,3,4;
'Zelle2=0: PosVal: 2,3,5;
'Zelle3=0: PosVal: 3,5;
'=> der Gesuchte Wert in Zelle1 ist 4
        If OptionSolve > 2 Then
          Call GetLonelyPosVals
        End If

        'Call GetMissingValInBLC
        'If OptionSolve > 0 Then
        '  Call GetMissingValInBLC
        '  Call GetPossibleValsToVals
        'End If

'        For Each V In ValueCol
'          V.UpdateToolTip
'        Next
'        wird in Show gemacht
'zum Schluß prüfen, ob schon fertig, nächste Iteration soll
'der User selber machen, durch nochmaliges klicken auf Solve
'oder Option Overall Iteration
  'Else ' GetAmountOfUnzeroVals = 81
    'DoTheCheck 'wird im Formular gemacht
      End If
    End If
  End If
End Sub
Private Sub GetMissingValInBLC()
Dim B As SudokuBLC
Dim L As SudokuBLC
Dim C As SudokuBLC
  'alle Blöcke durchwandern
  For Each B In mBlockCol
    B.FindMissingVals
  Next
  'alle Zeilen durchwandern
  For Each L In mLineCol
    L.FindMissingVals
  Next
  'alle Spalten durchwandern
  For Each C In mColmCol
    C.FindMissingVals
  Next
End Sub
Private Sub GetPossibleValsToVals()
Dim B As SudokuBLC 'Block
Dim L As SudokuBLC 'Line
Dim C As SudokuBLC 'Column
Dim V As SudokuVal
  For Each V In ValueCol
    'der Block in der sich die Zelle befindet
    Set B = mBlockCol(V.BlckNr)
    'die Zeile in der sich die Zelle befindet
    Set L = mLineCol(V.LineNr)
    'die Spalte in der sich die Zelle befindet
    Set C = mColmCol(V.ColmNr)
    Call V.GetPossibleVals(B.MissPossCol, L.MissPossCol, C.MissPossCol)
  Next
End Sub
Private Sub GetSinglePosValToCell()
Dim V As SudokuVal
  For Each V In ValueCol
    Call V.SetSinglePosVal
  Next
End Sub
Public Sub GetLonelyPosVals()
Dim B As SudokuBLC
Dim L As SudokuBLC
Dim C As SudokuBLC
  For Each B In mBlockCol
    B.FindLonelyPossibleVals
  Next
  'alle Zeilen durchwandern
  For Each L In mLineCol
    L.FindLonelyPossibleVals
  Next
  'alle Spalten durchwandern
  For Each C In mColmCol
    C.FindLonelyPossibleVals
  Next
End Sub

Private Function GetTryFillOutVal() As String
Dim V As SudokuVal, StrVal As String
Dim n As Long
  n = 1
  Do While (Len(StrVal) = 0) And (n < mn * mn)
    For Each V In ValueCol
      If V.CountMissPossVals = n Then
        StrVal = StrVal & V.ToString & vbCrLf
      End If
    Next
    n = n + 1
  Loop
  GetTryFillOutVal = StrVal
End Function
Public Sub DoTheCheck()
Dim mess As String
Dim n4 As Long, GAOUZV As Long
  n4 = mn * mn * mn * mn
  GAOUZV = GetAmountOfUnZeroVals
  If GAOUZV = n4 Then
    If Not CheckMe Then
      'mess = mess & "Das Sudoku-Rätsel konnte leider nicht richtig gelöst werden." & vbCrLf
      mess = mess & "Unfortunately the sudoku riddle could not be solved correctly." & vbCrLf
      'außerdem soll ausgegeben werden wo es gehakt hat:
      mess = mess & msCheckErr
      MsgBox mess
    Else
      'MsgBox "habe fertig"
      'MsgBox "Herzlichen Glückwunsch, sie haben das Sudoku richtig gelöst!", vbMsgBoxRight, "Sudoku Sudoku Sudoku"
      MsgBox "Congratulations, you have finally solved the sudoku correctly!", vbMsgBoxRight, "Sudoku Sudoku Sudoku"
    End If
'  Else
'    If Not CheckMe Then MsgBox msCheckErr
  End If
End Sub
Private Function CheckMe() As Boolean
Dim i As Long
Dim B As SudokuBLC
Dim L As SudokuBLC
Dim C As SudokuBLC
  msCheckErr = vbNullString
  CheckMe = True
  For i = 1 To mBlockCol.Count
    Set B = mBlockCol(i)
    If Not B.CheckMe Then
      'msCheckErr = msCheckErr & "Block " & CStr(i) & " meldet:"
      msCheckErr = msCheckErr & "Block " & CStr(i) & " report:"
      msCheckErr = msCheckErr & B.CheckErr & vbCrLf
      CheckMe = False
    End If
  Next
  For i = 1 To mLineCol.Count
    Set L = mLineCol(i)
    If Not L.CheckMe Then
      'msCheckErr = msCheckErr & "Line " & CStr(i) & " meldet:"
      msCheckErr = msCheckErr & "Line " & CStr(i) & " report:"
      msCheckErr = msCheckErr & L.CheckErr & vbCrLf
      CheckMe = False
    End If
  Next
  For i = 1 To mColmCol.Count
    Set C = mColmCol(i)
    If Not C.CheckMe Then
      'msCheckErr = msCheckErr & "Column " & CStr(i) & " meldet:"
      msCheckErr = msCheckErr & "Column " & CStr(i) & " report:"
      msCheckErr = msCheckErr & C.CheckErr & vbCrLf
      CheckMe = False
    End If
  Next
End Function
'Public Sub SaveCompleteUndo()
'Dim VN As SudokuVal, VO As SudokuVal
'Dim OldValCol As Collection
'Dim i As Long, n As Long
'  'wird von Solve aufgerufen
'  Set OldValCol = ValueCol
'  UndoList.Add OldValCol
'  Set ValueCol = New_VCV2BLC 'New Collection
'  n = OldValCol.Count
'  For i = 1 To n
'    Set VO = OldValCol(i)
'    Set VN = ValueCol(i)
'    Call VO.CopyTo(VN)
'  Next
'  Call InitVal2TxtB
'End Sub
Public Sub ReadFromFile(PFN As String)
'Dim mStrCol As New Collection
Dim FNr As Integer, StrLine As String
TryE: On Error GoTo Catch
  FNr = FreeFile
  Open PFN For Binary Access Read As #FNr
  StrLine = String(LOF(FNr), vbNullChar)
  'Strline = Input$(LOF(FNr), FNr)
  Get #FNr, , StrLine
  Close #FNr
  GameName = ExtractFileName(PFN)
  Call ReadNParseFromStr(StrLine)
  Exit Sub
Catch:
  Close #FNr
End Sub
Public Function ExtractFileName(PFN As String) As String
Dim RevPathSepPos As Long, RevExtSepPos As Long
  ExtractFileName = PFN
  RevPathSepPos = InStrRev(ExtractFileName, "\")
  If RevPathSepPos > 0 Then ExtractFileName = Right$(ExtractFileName, Len(ExtractFileName) - RevPathSepPos)
  RevExtSepPos = InStrRev(ExtractFileName, ".")
  If RevExtSepPos > 0 Then ExtractFileName = Left$(ExtractFileName, RevExtSepPos - 1)
End Function
Public Sub ReadNParseFromStr(StrLine As String)
  If Not Len(StrLine) = 0 Then
    RaiseEvent SaveOnSolve(Me)
    DoEvents
    Call Parse(StrLine)
  End If
End Sub
Private Sub Parse(StrLine As String)
'Parses and loads the game from the string
Dim i As Long, j As Long, C As String * 1
Dim V As SudokuVal, LV As Long, ascC As Long
  If InStr(1, StrLine, "GameLevel=") Then
    mGLevel = Mid$(StrLine, 11, 4)
    StrLine = Right$(StrLine, Len(StrLine) - 15)
  End If
  'alle nicht Nummern entfernen
  For i = Len(StrLine) To 1 Step -1
    C = Mid$(StrLine, i, 1)
    Select Case Asc(C)
    '    " " "." "0" "1" "2" "3" "4" "5" "6" "7" "8" "9"
    Case 32, 46, 48 To 57
    '    "A" "B" "C" "D" "E" "F" "G" "X" "a" "b" "c" "d" "e" "f" "g"
    Case 65 To 71, 88, 97 To 103
    Case Else 'alles andere rauslöschen durch überschreiben
      Mid(StrLine, i, 1) = " "
    End Select
  Next
  'jetzt die Zahlen auslesen, und den Zellen zuweisen
  For i = 1 To Len(StrLine)
    C = Mid$(StrLine, i, 1)
    ascC = Asc(C)
    If ascC <> 32 Then
      Select Case ascC
      Case 46: LV = 0
      Case 48 To 57:  LV = ascC - 48 'auch "0"
      Case 65 To 71:  LV = ascC - 55
      Case 88: LV = 0 '"0"
      Case 97 To 103: LV = ascC - 87
      End Select
      j = j + 1
      If j <= ValueCol.Count Then
        Set V = ValueCol(j)
        V.Value = LV
      End If
    End If
  Next
End Sub
Public Sub WriteToFile(PFN As String)
Dim Idx As Long, Offset As Long, jj As Long, ii As Long
'Dim L As SudokuBLC
Dim V As SudokuVal
Dim FNr As Integer
Dim h As Long, i As Long, j As Long, k As Long, n2 As Long
Dim StrVal As String, str As String
TryE: On Error GoTo Catch
  FNr = FreeFile
  n2 = mn * mn
  Open PFN For Output As #FNr
  StrVal = StrVal & "GameLevel=" & mGLevel & vbCrLf
  For k = 1 To mn
    For j = 1 To mn
      jj = j + (k - 1) * mn '3 'für Zeilen und Spalten
      For i = 1 To mn
        For h = 1 To mn
          ii = h + (i - 1) * mn '3 'für Zeilen und Spalten
          Idx = ii + (jj - 1) * n2 '9)
          Set V = ValueCol(Idx)
          If V > 9 Then str = Chr$(V + 55) Else str = CStr(V)
          StrVal = StrVal & str & " "
        Next
        If i < mn Then
          StrVal = StrVal & "| "
        Else
          StrVal = StrVal & vbCrLf
        End If
      Next
    Next
    If k < mn Then
      StrVal = StrVal & String$(2 * mn * mn + 2 * (mn - 1), "-")
    End If
    StrVal = StrVal & vbCrLf
  Next
  Print #FNr, StrVal
  'MsgBox "not yeti plemented " & PFN
Catch:
  Close #FNr
End Sub
