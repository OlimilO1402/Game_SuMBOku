VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SudokuVal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit '164 Zeilen
'stellt eine Zelle dar
'(nicht Zeile nein Zelllle)
'Public value As Long 'der Wert Attribute: Default
Private m_Value As Long
'sollte das nicht besser String sein? oder doch besser Byte
'Ok man könnte auch werte größer als 9 als string interpretieren
'also 10 = chr(65) = "A"
'if value > 9 then str = chr$(value+55) else str = cstr(value)

'Wozu brauchen wir eigentlich gleich wieder die ganzen Nummern?
Public Nr As Long 'die Nummer im Gesamtspielfeld
Public BlckNr As Long  'Blocknummer
Public LineNr As Long  'Zeilennummer
Public ColmNr As Long  'Spaltennummer
'speichert mögliche Werte als Long zw. 1-9
'Public PossibleVals As Collection
'Private WithEvents
Implements SudokuMissPoss
Private Base As New SudokuMissPoss
'
'umm gefährlich Daten und Ansicht hier nicht sauber getrennt
Private mTxtBox As TextBox
Private mTT As String 'ToolTipText

Private Sub Class_Initialize()
  'Call ClearPosVals 'schon in Base New erledigt
End Sub
Friend Sub NewCC(other As SudokuVal)
'Copyconstructor geht eigentlich nur richtig wenn private/protected Zugriff auf Member besteht
'ehemals Sub CopyTo(Dest as SudokuVal)
Dim i As Long, PV As Long, n As Long
  'beim Undo nur Werte kopieren, alles was durch Funktionen ermittelt
  'werden kann erst dann ermitteln, wenn auf Undo geklickt wird.
  'das ist hier eigentlich nur Solve, bzw. GetMissingVals.
  'halt Neeiiiiiiiin genau andersherum
  'hier übernehme ich Wert vom anderen Object
  Value = other.Value
  Nr = other.Nr
  Me.BlckNr = other.BlckNr
  Me.LineNr = other.LineNr
  Me.ColmNr = other.ColmNr
  Set Base = New_SudokuMissPoss(other)
  'Dest = Me.value
  'Dest.Nr = Me.Nr
  'Dest.BlckNr = Me.BlckNr
  'Dest.ColmNr = Me.ColmNr
  'Dest.LineNr = Me.LineNr
  'Set Dest.TxtBox = Me.TxtBox
  'n = CountMissPossVals
  'For i = 1 To n
  '  PV = MissPossItem(i) 'PossibleVals(i)
  '  Call Dest.AddMissPossVal(PV)
  'Next
End Sub
Public Property Get Value() As Long
    Value = m_Value
End Property
Public Property Let Value(v As Long)
    m_Value = v
    'jetzt den wert aus der missposs rausnehmen?
    Call Base.DeleteMissPossVal(v)
End Property
Public Sub CopyTo(Dest As SudokuVal)
  Dest = Me.Value
  Dest.Nr = Me.Nr
  Dest.BlckNr = Me.BlckNr
  Dest.ColmNr = Me.ColmNr
  Dest.LineNr = Me.LineNr
  Set Dest.TxtBox = Me.TxtBox
  n = CountMissPossVals
  For i = 1 To n
    PV = MissPossItem(i) 'PossibleVals(i)
    Call Dest.AddMissPossVal(PV)
  Next
End Sub
Public Function ToTextBox()
Dim str As String
  If Not mTxtBox Is Nothing Then
    If Value > 0 Then
      If Value > 9 Then str = Chr$(Value + 55) Else str = CStr(Value)
      mTxtBox.Text = str
    Else
      mTxtBox.Text = vbNullString
    End If
  End If
End Function
Public Function ToString() As String
  If Len(mTT) = 0 Then mTT = "Nr:" & CStr(Nr) & " B:" & CStr(BlckNr) & " L:" & CStr(LineNr) & " C:" & CStr(ColmNr)
  ToString = mTT & " PossV: " & MissPossToString
End Function
Public Property Get TxtBox() As TextBox
  Set TxtBox = mTxtBox
End Property
Public Property Set TxtBox(ByVal TBVal As TextBox)
  Set mTxtBox = TBVal
  mTT = "Nr:" & CStr(Nr) & " B:" & CStr(BlckNr) & " L:" & CStr(LineNr) & " C:" & CStr(ColmNr)
End Property
Public Sub UpdateToolTip()
  If Not mTxtBox Is Nothing Then
    mTxtBox.ToolTipText = ToString
  End If
End Sub
Public Sub Clear()
  Value = 0
  If Not mTxtBox Is Nothing Then
    mTxtBox.Text = vbNullString '"" 'vbNullString
    mTxtBox.ToolTipText = mTT
  End If
  ClearMissPossCol 'ClearPosVals
End Sub
'Public Sub RemovePosVal(BVal As Long)
'Dim i As Integer, n As Long, B As Long
'  n = CountMissPossVals 'PossibleVals.Count
'  If n > 0 Then
'    For i = n To 1 Step -1
'      B = PossibleVals(i)
'      If B = BVal Then PossibleVals.Remove (i)
'    Next
'  End If
'End Sub
'nur die Werte dazunehmen die in allen drei Collections enthalten sind
Public Sub GetPossibleVals(Col1 As Collection, Col2 As Collection, Col3 As Collection)
Dim bi As Long, bj As Long, bk As Long
Dim i As Long, j As Long, k As Long
  'Call ClearPosVals
  Call ClearMissPossCol
  If Value = 0 Then
    For i = 1 To Col1.Count
      bi = Col1(i)
      For j = 1 To Col2.Count
        bj = Col2(j)
        For k = 1 To Col3.Count
          bk = Col3(k)
          If bi = bj And bj = bk Then
            AddMissPossVal (bi)
          End If
        Next
      Next
    Next
  End If
End Sub
Public Sub SetSinglePosVal()
  'If PossibleVals.Count = 1 Then
  If CountMissPossVals = 1 Then
    Value = MissPossItem(1)
  End If
End Sub

'######################'   Implements SudokuMissPoss   '######################'
Private Sub SudokuMissPoss_AddMissPossVal(LngVal As Long)
    Call AddMissPossVal(LngVal)
End Sub
Public Sub AddMissPossVal(LngVal As Long)
    Call Base.AddMissPossVal(LngVal)
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
Public Sub ClearMissPossCol() 'ClearMisVals
  'Set MissingVals = New Collection
  Call Base.ClearMissPossCol
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
Private Function MissPossToString() As String 'PosValsToString
  MissPossToString = Base.ToString
End Function
