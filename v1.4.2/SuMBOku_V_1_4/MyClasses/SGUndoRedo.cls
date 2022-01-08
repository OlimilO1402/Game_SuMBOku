VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SGUndoRedo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit '107 Zeilen
'eine klasse die für das Game die Aufgaben des UndoRedo übernimmt
'also wenn der UndoButton geklickt wird, dann soll diese Klasse
'das undo vornehmen, vorher soll diese Klasse die Werte übernehmen
'Als erstes sollte die Frage geklärt werden in welcher Beziehung
'das SudokuGame und die UndoKlasse zueinander stehen
'Gleichberechtigt nebeneinander im Hauptprogramm, bzw im Formular
'
'Undo und Redo für SudokuGame
Private UndoList As Collection 'speichert Undoinstanzen von ValueCol (des ganzen Games)
Private RedoList As Collection 'speichert RedoInstanzen des ganzen Games
Private WithEvents mBtnUndo As CommandButton
Attribute mBtnUndo.VB_VarHelpID = -1
Private WithEvents mBtnRedo As CommandButton
Attribute mBtnRedo.VB_VarHelpID = -1
Public WithEvents Sudoku As SudokuGame
Attribute Sudoku.VB_VarHelpID = -1

Private Sub Class_Initialize()
  Call ClearUndoList
  Call ClearRedoList
End Sub

Public Sub NewC(aSudokuGame As SudokuGame)
   Set Sudoku = aSudokuGame
End Sub

Public Sub ClearUndoList()
  Set UndoList = New Collection
End Sub
Public Sub ClearRedoList()
  Set RedoList = New Collection
End Sub

Public Sub SetBtnUndoRedo(BtnUndo As CommandButton, BtnRedo As CommandButton)
  Set mBtnUndo = BtnUndo
  Set mBtnRedo = BtnRedo
  Call EnDisAbleButtons
End Sub

Public Sub SaveCompleteUndo()
'   VN = new;        VO = old
Dim VN As SudokuVal, VO As SudokuVal
Dim OldValCol As Collection
Dim i As Long, n As Long
  Set OldValCol = Sudoku.ValueCol
  UndoList.Add OldValCol
  Set Sudoku.ValueCol = Sudoku.New_VCV2BLC 'New Collection
  n = OldValCol.Count
  For i = 1 To n
    Set VO = OldValCol(i)
    Set VN = Sudoku.ValueCol(i)
    'Call VO.CopyTo(VN)
    '
    Set VN = New_SudokuVal(VO)
    
  Next
  Call Sudoku.InitVal2TxtB
  Call EnDisAbleButtons
End Sub

Public Sub UndoLastAction()
Dim i As Long, n As Long
Dim o As Object
  n = UndoList.Count
  If n > 0 Then
    Set o = UndoList(n)
    If TypeOf o Is Collection Then
      'das ganze Spielfeld rückgängig machen
      RedoList.Add Sudoku.ValueCol 'in die Redoliste eintragen
      Set Sudoku.ValueCol = UndoList(n) 'den alten Wert aus der Undoliste zuweisen
      Call UndoList.Remove(n)  'aus der Undoliste löschen
    Else
      'ein einzelner Wert
      If TypeOf o Is SudokuVal Then
        'einen einzelnen Wert rückgängig machen
        
      End If
    End If
  End If
  Call EnDisAbleButtons
End Sub
Public Sub RedoLastAction()
Dim i As Long, n As Long
Dim o As Object
  n = RedoList.Count
  If n > 0 Then
    Set o = RedoList(n)
    If TypeOf o Is Collection Then
      'das ganze Spielfeld wiederherstellen
      UndoList.Add Sudoku.ValueCol 'in die Redolsite eintragen
      Set Sudoku.ValueCol = RedoList(n) 'den alten Wert aus der Undoliste zuweisen
      Call RedoList.Remove(n)  'aus der Undoliste löschen
    Else
      'ein einzelner Wert
      If TypeOf o Is SudokuVal Then
        'einen einzelnen Wert wiederherstellen
      
      End If
    End If
  End If
  Call EnDisAbleButtons
End Sub
Private Sub EnDisAbleButtons()
  If RedoList.Count = 0 Then
    mBtnRedo.Enabled = False
  Else
    mBtnRedo.Enabled = True
  End If
  If UndoList.Count = 0 Then
    mBtnUndo.Enabled = False
  Else
    mBtnUndo.Enabled = True
  End If
End Sub

Private Sub Sudoku_SaveOnSolve(Obj As SudokuGame)
  Call SaveCompleteUndo
End Sub
