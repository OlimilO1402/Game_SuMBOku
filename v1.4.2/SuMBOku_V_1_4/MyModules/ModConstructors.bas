Attribute VB_Name = "ModConstructors"
Option Explicit

Public Function New_Font(aName As String, aSize As Long) As StdFont
  Set New_Font = New StdFont
  New_Font.Name = aName
  New_Font.Size = aSize * 72 / 96 '/ 72 'Screen.TwipsPerPixelX
End Function
Public Function New_DynTextBoxes(ParentForm As Form) As DynTextBoxes
  Set New_DynTextBoxes = New DynTextBoxes
  Call New_DynTextBoxes.NewC(ParentForm)
End Function

Public Function New_DynTextBox(aFrm As Form, aTBCol As DynTextBoxes, aStrName As String, index As Long) As DynTextBox
  Set New_DynTextBox = New DynTextBox
  Call New_DynTextBox.NewC(aFrm, aTBCol, aStrName, index)
End Function

Public Function New_Image(aPic As StdPicture) As Image
  Set New_Image = New Image
  Call New_Image.FromStdPicture(aPic)
End Function
Public Function LoadFromResource(nID As Long) As IPictureDisp
  Set LoadFromResource = LoadResPicture(nID, vbResBitmap)
End Function

Public Function New_SudokuGame(ByVal sizemn As Long) As SudokuGame
   Set New_SudokuGame = New SudokuGame
   Call New_SudokuGame.NewC(sizemn)
End Function

Public Function New_SGUndoRedo(aSudokuGame As SudokuGame) As SGUndoRedo
    Set New_SGUndoRedo = New SGUndoRedo
    'Set New_SGUndoRedo.Sudoku = aSudokuGame
    Call New_SGUndoRedo.NewC(aSudokuGame)
End Function

Public Function New_SudokuVal(other As SudokuVal) As SudokuVal
   'CopyConstructor
   Set New_SudokuVal = New SudokuVal
   Call New_SudokuVal.NewCC(other)
End Function

Public Function New_SudokuMissPoss(other As SudokuMissPoss) As SudokuMissPoss
   'CopyConstructor
   Set New_SudokuMissPoss = New SudokuMissPoss
   Call New_SudokuMissPoss.NewCC(other)
End Function


'ist Shared muﬂ deshalb in ein Modul
'von Bitmap:
'Public Shared Function FromResource(ByVal hinstance As System.IntPtr, ByVal bitmapName As String) As System.Drawing.Bitmap
Public Function FromResource(Optional ByVal hhinstance As Long, Optional ByVal bitmapNameResID As String) As Image 'System.Drawing.Bitmap
  Set FromResource = New Image
  Call FromResource.FromStdPicture(LoadResPicture(CLng(bitmapNameResID), vbResBitmap))
End Function


