Attribute VB_Name = "MMain"
Option Explicit '9 Zeilen
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Sub Main()
    Call InitCommonControls
    FrmSudoku.Show
    FrmSudoku.SetupGame
End Sub

'in der Klasse SudokuGame
'nicht mehr gebrauchte Sub:
' Private Sub InitVal2BLCs(ValCol As Collection)
'  'jetzt noch die einzelnen SudokuVal in den
'  'Blöcken, Zeilen und Spalten speichern
'  For j = 1 To n2 '9
'    Set L = mLineCol(j)
'    Set C = mColmCol(j)
'    For i = 1 To n2 '9
'      'in die Zeilen speichern
'      Set V = ValueCol(i + (j - 1) * n2) '9)
'      L.mValCol.Add V
'      'in die Spalten speichern
'      Set V = ValueCol(j + (i - 1) * n2) '9)
'      C.mValCol.Add V
'    Next
'  Next
  'in die Blöcke speichern is bissl kniffliger deshalb hier mal getrennt
'  For k = 1 To mn
'    For j = 1 To mn
'      Set B = mBlockCol(jj)
'      Offset = ((k - 1) * n3) + ((j - 1) * mn) 'für die Blöcke
'      For i = 1 To mn
'        For h = 1 To mn
'          'in die Blöcke
'          Idx = (h + (i - 1) * n2) + Offset
'          Set V = ValueCol(Idx)
'          B.mValCol.Add V
'        Next
'      Next
'    Next
'  Next
'End Sub

