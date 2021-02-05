VERSION 5.00
Begin VB.Form frmSuEdit 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Line 1:"
   ClientHeight    =   285
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   285
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "1, 2, 3, 4, 5, 6, 7, 8, 9, A, B, C, D, E, F, G"
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "frmSuEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '19 Zeilen
Private mBLCV As SudokuMissPoss 'SudokuBLC

Public Property Set BLCV(aBLCV As SudokuMissPoss)
  Set mBLCV = aBLCV
  If Not mBLCV Is Nothing Then
    Text1.Text = mBLCV.ToString '.StrMissingVals
  End If
End Property
Private Sub Form_Unload(Cancel As Integer)
  mBLCV.Parse (Text1.Text)
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    Unload Me
  End If
End Sub
