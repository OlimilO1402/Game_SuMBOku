VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SudokuMissPoss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit '65 Zeilen
'stellt die Basisklasse dar für SudokuBLC und SudokuVal
'in block, line, column Missing values
'in cells Possible values
Private mCol As Collection 'speichert Werte vom Typ Long

Private Sub Class_Initialize()
  ClearMissPossCol
End Sub
Friend Sub NewCC(other As SudokuMissPoss)
   Dim v 'As Long
   'Dim othCol As Collection: Set othCol = other.MissPossCol
   For Each v In mCol
      'othCol.Add v
      other.MissPossCol.Add v
   Next
End Sub
Public Sub AddMissPossVal(LngVal As Long)
  mCol.Add LngVal
End Sub
Private Function mColContains(LngVal As Long) As Long
    Dim i As Long
    mColContains = 0 '-1 'Collection ist 1-basiert
    For i = 1 To mCol.Count
        If mCol.Item(i) = LngVal Then
            mColContains = i
            Exit For
        End If
        'mCol.Remove
    Next
End Function
Public Sub DeleteMissPossVal(LngVal As Long)
    Dim i As Long: i = mColContains(LngVal)
    If i > 0 Then Call mCol.Remove(i)
End Sub
Public Sub ClearMissPossCol()
  Set mCol = New Collection
End Sub
 
Public Function CountMissPossVals() As Long
  CountMissPossVals = mCol.Count
End Function

Public Function MissPossItem(index As Long) As Long
  MissPossItem = mCol(index)
End Function

Public Property Get MissPossCol() As Collection
  Set MissPossCol = mCol
End Property
Public Function ToString() As String
Dim bb As Long, i As Long, n As Long, StrV As String
Dim str As String
  'If Value = 0 Then 'Wieso eigentlich?
    n = mCol.Count
    If n > 0 Then
      For i = 1 To n
        bb = mCol(i)
        If bb > 9 Then str = Chr$(bb + 55) Else str = CStr(bb)
        StrV = StrV & str
        If i < mCol.Count Then StrV = StrV & ", "
      Next
      ToString = StrV
    End If
  'End If
End Function

Public Sub Parse(ByVal StrVals As String)
Dim StrVal As String * 1
Dim PV As Long, i As Long
    ClearMissPossCol
    For i = 1 To Len(StrVals)
      StrVal = Mid$(StrVals, i, 1)
      PV = Asc(Left(StrVal, 1))
      Select Case PV
      Case 49 To 57
        PV = PV - 48 '1 .. 9
        Call mCol.Add(PV)
      Case 65 To 72
        PV = PV - 55 'A .. G = 10 .. 16
        Call mCol.Add(PV)
      Case Else
      End Select
    Next
End Sub
