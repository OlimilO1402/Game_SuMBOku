VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RandomGames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit '26 Zeilen
Private mCol As New Collection
Public GameName As String
'Public GamLevel As Long

Private Sub Class_Initialize()
  '
End Sub
Private Sub LoadGames()
Dim i As Long
  For i = 1001 To 1060
    mCol.Add (LoadResString(i))
  Next
End Sub
'GameLevel=eins 870001400300204000000000009001006094900307008280400600500000000000109003004500027 very_easy1
Public Function GetNextGame() As String
Dim i As Long
  Randomize
  If mCol.Count = 0 Then LoadGames
    
  If mCol.Count > 0 Then
    i = Int(mCol.Count * Rnd + 1)
    If i = 0 Then i = 1
    GetNextGame = mCol(i)
    mCol.Remove (i)
    GameName = Right$(GetNextGame, Len(GetNextGame) - 97)
    GetNextGame = Left$(GetNextGame, 98)
  End If
End Function
