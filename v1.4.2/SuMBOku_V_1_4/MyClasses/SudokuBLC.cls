VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SudokuBLC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit '219 Zeilen
'stellt Block, Zeile, Spalte eines Spielfeldes dar
'<Implements> und <Private Base> deckt die ganze Palette für
'Funktionalität der Vererbung ab. was will man mehr...
Implements SudokuMissPoss
Private Base As New SudokuMissPoss 'Liste speichert fehlende Werte
Public Nr As Long
Public mn As Long
Public ValCol As Collection 'speichert Verweise auf 9 Zellen
Public CheckErr As String   'das Ergebnis von CheckMe
Private mImg As VB.Image         'das Feld für den ToolTip
Private mBLCTyp As SuBLCTyp   'was bin ich?
Private mStrTyp As String
Public Enum SuBLCTyp
  stBlock = 66 '=Asc("B")
  stLine = 76  '=Asc("L")
  stColum = 67 '=Asc("C")
End Enum

Private Sub Class_Initialize()
  Call Clear
  'Call ClearMisVals ClearMissPossCol 'wurde schon in Base New gemacht
End Sub
Public Sub Clear()
  Set ValCol = New Collection
End Sub
Public Sub SetVal(aVal As Long)
Dim v As SudokuVal
  For Each v In ValCol
    v.Value = aVal
  Next
End Sub
Public Property Let BLCTyp(LngVal As SuBLCTyp)
  mBLCTyp = LngVal
  Select Case mBLCTyp
  Case stBlock: mStrTyp = "Block"
  Case stLine: mStrTyp = "Line"
  Case stColum: mStrTyp = "Column"
  End Select
End Property
Public Property Set img(ImgVal As VB.Image)
  Set mImg = ImgVal
  UpdateToolTip
End Property
Public Sub UpdateToolTip()
Dim str As String
  str = MissPossToString
  If Len(str) > 0 Then str = " MissV: " & str
  If Not mImg Is Nothing Then mImg.ToolTipText = mStrTyp & CStr(Nr) & str
End Sub

Public Sub FindMissingVals()
Dim bb As Long, IsIn As Boolean
Dim v As SudokuVal, n2 As Long
  ClearMissPossCol
  n2 = mn * mn
  For bb = 1 To n2
    For Each v In ValCol
      If v.Value > 0 Then
        If v.Value = bb Then
          IsIn = True
          Exit For
        End If
      End If
    Next
    If Not IsIn Then
      AddMissPossVal (bb)
    Else
      IsIn = False
    End If
  Next
  UpdateToolTip
End Sub

'Private Sub RemoveVal(Val As Long)
'Dim i As Long, V As Long
'  For i = 1 To mMissingVals.Count
'    V = mMissingVals(i)
'    If V = Val Then
'      mMissingVals.Remove (i)
'      Exit Function
'    End If
'  Next
'End Sub

'wenn alle Werte größer Null sind, soll CheckMe untersuchen
'ob jeder Wert nur einmal vorkommt, bzw. wenn nicht dann soll
'ausgegeben werden welche Werte mehr als einmal vorkommen und
'welche garnicht.
Public Function CheckMe() As Boolean
Dim CAlle As New Collection 'speichert Collections die eizelne Werte speichern
Dim C1To9 As New Collection 'speichert einzelne Werte
Dim i As Long, v As SudokuVal, n2 As Long, NUZV As Long, n As Long
  CheckErr = vbNullString
  n2 = mn * mn
  NUZV = GetAmountOfUnZeroVals 'NumberofUnZeroVals
  n = ValCol.Count
'nur was machen wenn alle Zellen einen Eintrag haben
  If NUZV = n Then
    For i = 1 To n2 '9
      Set C1To9 = New Collection
      CAlle.Add C1To9
    Next
    For Each v In ValCol
      If v.Value > 0 Then
        Set C1To9 = CAlle(v.Value)
        C1To9.Add v
      End If
    Next
    For i = 1 To CAlle.Count
      Set C1To9 = CAlle(i)
      If C1To9.Count <> 1 Then
        If C1To9.Count = 0 Then
          CheckErr = CheckErr & "  Wert: " & CStr(i) & " kommt garnicht vor; "
        Else
          CheckErr = CheckErr & "  Wert: " & CStr(i) & " kommt: " & CStr(C1To9.Count) & "-mal vor; "
        End If
      End If
    Next
    'If Len(CheckErr) > 0 Then CheckErr = CheckErr & vbCrLf
    If Len(CheckErr) = 0 Then CheckMe = True
  Else
    CheckErr = CStr(n - NUZV) & ": missing values"
  End If
End Function

Public Function GetAmountOfUnZeroVals() As Long
Dim v As SudokuVal
  For Each v In ValCol
    If v > 0 Then GetAmountOfUnZeroVals = GetAmountOfUnZeroVals + 1
  Next
End Function

'die Werte finden die in den PosVals nur einmal vorkommen
Public Sub FindLonelyPossibleVals()
Dim v As SudokuVal, i As Long, bb As Long, n As Long, n2 As Long
Dim CAlle As New Collection
Dim C1To9 As Collection
Dim StrVal As String
  n2 = mn * mn
  For i = 1 To n2 '9
    Set C1To9 = New Collection
    CAlle.Add C1To9
  Next
  For Each v In ValCol
    n = v.CountMissPossVals 'PossibleVals.Count
    For i = 1 To n
      bb = v.MissPossItem(i)
      Set C1To9 = CAlle(bb)
      C1To9.Add v
    Next
  Next
  n = CAlle.Count
  For i = 1 To n
    Set C1To9 = CAlle(i)
    If C1To9.Count = 1 Then
      Set v = C1To9(1)
      v = CLng(i)
      'MsgBox "in Block: " & CStr(V.BlckNr) & " die Zahl: " & CStr(i) & " im Feld: " & CStr(V.Nr)
      'den Wert in allen possibleVals entfernen
      'For Each V In ValCol
      '  Call V.RemovePosVal(i)
      'Next
    End If
  Next
End Sub

'######################'   Implements SudokuMissPoss   '######################'
Private Sub SudokuMissPoss_AddMissPossVal(LngVal As Long)
  Call AddMissPossVal(LngVal)
End Sub
Public Sub AddMissPossVal(LngVal As Long)
  Base.AddMissPossVal (LngVal)
End Sub

Public Sub SudokuMissPoss_DeleteMissPossVal(LngVal As Long)
    Call DeleteMissPossVal(LngVal)
End Sub
Private Sub DeleteMissPossVal(LngVal As Long)
    Call Base.DeleteMissPossVal(LngVal)
End Sub


Private Sub SudokuMissPoss_ClearMissPossCol()
  Call ClearMissPossCol
End Sub
Public Sub ClearMissPossCol()
  Call Base.ClearMissPossCol
  Call UpdateToolTip
End Sub

Private Function SudokuMissPoss_CountMissPossVals() As Long
  SudokuMissPoss_CountMissPossVals = CountMissPossVals
End Function
Public Function CountMissPossVals() As Long
  CountMissPossVals = Base.CountMissPossVals
End Function

Private Property Get SudokuMissPoss_MissPossCol() As Collection
  Set SudokuMissPoss_MissPossCol = MissPossCol
End Property
Public Property Get MissPossCol() As Collection
  Set MissPossCol = Base.MissPossCol
End Property

Private Function SudokuMissPoss_MissPossItem(index As Long) As Long
  SudokuMissPoss_MissPossItem = MissPossItem(index)
End Function
Public Function MissPossItem(index As Long) As Long
  MissPossItem = Base.MissPossItem(index)
End Function

Private Sub SudokuMissPoss_Parse(ByVal StrVals As String)
  Call ParseMissPossStr(StrVals)
End Sub
Private Sub ParseMissPossStr(StrVals As String)
  Call Base.Parse(StrVals)
  UpdateToolTip
End Sub

Private Function SudokuMissPoss_ToString() As String
  SudokuMissPoss_ToString = MissPossToString
End Function
Private Function MissPossToString() As String
  MissPossToString = Base.ToString
End Function

