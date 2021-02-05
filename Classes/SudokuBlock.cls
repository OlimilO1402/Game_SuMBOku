VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SudokuBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'stellt einen Block dar
Public Nr As Long
Public mn As Long
Public ValCol As Collection 'speichert Verweise auf 9 Zellen
Public MissingVals As Collection
Public CheckErr As String 'das Ergebnis von CheckMe

Private Sub Class_Initialize()
  Call Clear
  Call ClearMisVals
End Sub
Public Sub Clear()
  Set ValCol = New Collection
End Sub
Private Sub ClearMisVals()
  Set MissingVals = New Collection
End Sub
Public Sub SetVal(aVal As Long)
Dim V As SudokuVal
  For Each V In ValCol
    V.Value = aVal
  Next
End Sub

Public Sub GetMissingVals()
Dim bb As Long, IsIn As Boolean
Dim V As SudokuVal, n2 As Long
  Call ClearMisVals
  n2 = mn * mn
  For bb = 1 To n2
    For Each V In ValCol
      If V.Value > 0 Then
        If V.Value = bb Then
          IsIn = True
          Exit For
        End If
      End If
    Next
    If Not IsIn Then
      MissingVals.Add bb
    Else
      IsIn = False
    End If
  Next
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
Dim i As Long, V As SudokuVal, n2 As Long, NUZV As Long, n As Long
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
    For Each V In ValCol
      If V.Value > 0 Then
        Set C1To9 = CAlle(V.Value)
        C1To9.Add V
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
Dim V As SudokuVal
  For Each V In ValCol
    If V.Value > 0 Then GetAmountOfUnZeroVals = GetAmountOfUnZeroVals + 1
  Next
End Function

'die Werte finden die in den PosVals nur einmal vorkommen
Public Sub GetLonelyPossibleVals()
Dim V As SudokuVal, i As Long, bb As Long, n As Long, n2 As Long
Dim CAlle As New Collection
Dim C1To9 As Collection
Dim StrVal As String
  n2 = mn * mn
  For i = 1 To n2 '9
    Set C1To9 = New Collection
    CAlle.Add C1To9
  Next
  For Each V In ValCol
    n = V.PossibleVals.Count
    For i = 1 To n
      bb = V.PossibleVals(i)
      Set C1To9 = CAlle(bb)
      C1To9.Add V
    Next
  Next
  n = CAlle.Count
  For i = 1 To n
    Set C1To9 = CAlle(i)
    If C1To9.Count = 1 Then
      Set V = C1To9(1)
      V.Value = CLng(i)
      'MsgBox "in Block: " & CStr(V.BlckNr) & " die Zahl: " & CStr(i) & " im Feld: " & CStr(V.Nr)
      'den Wert in allen possibleVals entfernen
      'For Each V In ValCol
      '  Call V.RemovePosVal(i)
      'Next
    End If
  Next
End Sub
