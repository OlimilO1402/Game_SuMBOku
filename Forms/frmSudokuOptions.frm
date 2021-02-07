VERSION 5.00
Begin VB.Form frmSudokuOptions 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Solve Options"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   3135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnCancel 
      Caption         =   "Abbrechen"
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Solve with..."
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'Kein
         Height          =   1095
         Left            =   120
         ScaleHeight     =   1095
         ScaleWidth      =   2655
         TabIndex        =   1
         Top             =   240
         Width           =   2655
         Begin VB.OptionButton Option4 
            Caption         =   "the whole iteration"
            Height          =   255
            Left            =   0
            TabIndex        =   7
            ToolTipText     =   "not yet implemented"
            Top             =   720
            Width           =   2535
         End
         Begin VB.OptionButton Option3 
            Caption         =   "all three steps"
            Height          =   255
            Left            =   0
            TabIndex        =   4
            Top             =   480
            Width           =   2535
         End
         Begin VB.OptionButton Option2 
            Caption         =   "the first and second step"
            Height          =   255
            Left            =   0
            TabIndex        =   3
            ToolTipText     =   "Second: fill in single possible values"
            Top             =   240
            Width           =   2535
         End
         Begin VB.OptionButton Option1 
            Caption         =   "only the first step"
            Height          =   255
            Left            =   0
            TabIndex        =   2
            ToolTipText     =   "No new values only possible values in ToolTips"
            Top             =   0
            Width           =   1935
         End
      End
   End
End
Attribute VB_Name = "frmSudokuOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '28 Zeilen
Public Sudoku As SudokuGame

Private Sub Form_Load()
  Option1.Value = False
  Option2.Value = False
  Option3.Value = False
End Sub
Private Sub Form_Activate()
  Select Case Sudoku.OptionSolve
  Case 1: Option1.Value = True
  Case 2: Option2.Value = True
  Case 3: Option3.Value = True
  Case 4: Option4.Value = True
  End Select
End Sub

Private Sub BtnOK_Click()
  If Option1.Value Then Sudoku.OptionSolve = 1
  If Option2.Value Then Sudoku.OptionSolve = 2
  If Option3.Value Then Sudoku.OptionSolve = 3
  If Option4.Value Then Sudoku.OptionSolve = 4
  Unload Me
End Sub
Private Sub BtnCancel_Click()
  Unload Me
End Sub
