VERSION 5.00
Begin VB.Form FrmSudoku 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "SuMBOku 1.2"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4815
   Icon            =   "FrmSudoku.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   401
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   321
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnOptions 
      Caption         =   "Options"
      Height          =   375
      Left            =   3360
      TabIndex        =   89
      ToolTipText     =   "Solve options"
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton BtnInfo 
      Caption         =   "Info"
      Height          =   375
      Left            =   2880
      TabIndex        =   94
      ToolTipText     =   "Information & help"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton RndExample 
      Caption         =   "Rnd Example"
      Height          =   375
      Left            =   1800
      TabIndex        =   95
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton BtnSave 
      Caption         =   "S"
      Height          =   375
      Left            =   3600
      TabIndex        =   93
      ToolTipText     =   "Save SudokuFile"
      Top             =   5520
      Width           =   375
   End
   Begin VB.CommandButton BtnOpen 
      Caption         =   "O"
      Height          =   375
      Left            =   3240
      TabIndex        =   92
      ToolTipText     =   "Open SudokuFile"
      Top             =   5520
      Width           =   375
   End
   Begin VB.CommandButton BtnExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   4080
      TabIndex        =   84
      ToolTipText     =   "Quit the game"
      Top             =   5520
      Width           =   615
   End
   Begin VB.CommandButton BtnRedo 
      Caption         =   "R"
      Height          =   375
      Left            =   1200
      TabIndex        =   91
      ToolTipText     =   "Redo last action"
      Top             =   5520
      Width           =   375
   End
   Begin VB.CommandButton BtnUndo 
      Caption         =   "U"
      Height          =   375
      Left            =   840
      TabIndex        =   90
      ToolTipText     =   "Undo last action"
      Top             =   5520
      Width           =   375
   End
   Begin VB.CommandButton BtnSolve 
      Caption         =   "Solve"
      Height          =   375
      Left            =   1680
      TabIndex        =   82
      ToolTipText     =   "do the next Solve-step "
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton BtnTestBlocks 
      Caption         =   "Blocks"
      Height          =   375
      Left            =   2520
      TabIndex        =   87
      Top             =   5520
      Width           =   615
   End
   Begin VB.CommandButton BtnTestColms 
      Caption         =   "Colms"
      Height          =   375
      Left            =   2520
      TabIndex        =   86
      Top             =   5520
      Width           =   615
   End
   Begin VB.CommandButton BtnTestLines 
      Caption         =   "Lines"
      Height          =   375
      Left            =   1920
      TabIndex        =   85
      Top             =   5520
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Level:"
      Height          =   615
      Left            =   0
      TabIndex        =   83
      Top             =   0
      Width           =   1695
      Begin VB.Image ImGL4 
         Height          =   375
         Left            =   1200
         Top             =   195
         Width           =   375
      End
      Begin VB.Image ImGL3 
         Height          =   375
         Left            =   840
         Top             =   195
         Width           =   375
      End
      Begin VB.Image ImGL2 
         Height          =   375
         Left            =   480
         Top             =   195
         Width           =   375
      End
      Begin VB.Image ImGL1 
         Height          =   375
         Left            =   120
         Top             =   195
         Width           =   375
      End
      Begin VB.Image ImGL0 
         Height          =   495
         Left            =   0
         Top             =   120
         Width           =   135
      End
   End
   Begin VB.CommandButton BtnClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   120
      TabIndex        =   81
      ToolTipText     =   "Clear the whole game"
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   81
      Left            =   4200
      TabIndex        =   80
      Text            =   "1"
      Top             =   4920
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   80
      Left            =   3720
      TabIndex        =   79
      Text            =   "1"
      Top             =   4920
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   79
      Left            =   3240
      TabIndex        =   78
      Text            =   "1"
      Top             =   4920
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   78
      Left            =   2640
      TabIndex        =   77
      Text            =   "1"
      Top             =   4920
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   77
      Left            =   2160
      TabIndex        =   76
      Text            =   "1"
      Top             =   4920
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   76
      Left            =   1680
      TabIndex        =   75
      Text            =   "1"
      Top             =   4920
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   75
      Left            =   1080
      TabIndex        =   74
      Text            =   "1"
      Top             =   4920
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   74
      Left            =   600
      TabIndex        =   73
      Text            =   "1"
      Top             =   4920
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   73
      Left            =   120
      TabIndex        =   72
      Text            =   "1"
      Top             =   4920
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   72
      Left            =   4200
      TabIndex        =   71
      Text            =   "1"
      Top             =   4440
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   71
      Left            =   3720
      TabIndex        =   70
      Text            =   "1"
      Top             =   4440
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   70
      Left            =   3240
      TabIndex        =   69
      Text            =   "1"
      Top             =   4440
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   69
      Left            =   2640
      TabIndex        =   68
      Text            =   "1"
      Top             =   4440
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   68
      Left            =   2160
      TabIndex        =   67
      Text            =   "1"
      Top             =   4440
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   67
      Left            =   1680
      TabIndex        =   66
      Text            =   "1"
      Top             =   4440
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   66
      Left            =   1080
      TabIndex        =   65
      Text            =   "1"
      Top             =   4440
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   65
      Left            =   600
      TabIndex        =   64
      Text            =   "1"
      Top             =   4440
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   64
      Left            =   120
      TabIndex        =   63
      Text            =   "1"
      Top             =   4440
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   63
      Left            =   4200
      TabIndex        =   62
      Text            =   "1"
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   62
      Left            =   3720
      TabIndex        =   61
      Text            =   "1"
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   61
      Left            =   3240
      TabIndex        =   60
      Text            =   "1"
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   60
      Left            =   2640
      TabIndex        =   59
      Text            =   "1"
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   59
      Left            =   2160
      TabIndex        =   58
      Text            =   "1"
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   58
      Left            =   1680
      TabIndex        =   57
      Text            =   "1"
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   57
      Left            =   1080
      TabIndex        =   56
      Text            =   "1"
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   56
      Left            =   600
      TabIndex        =   55
      Text            =   "1"
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   55
      Left            =   120
      TabIndex        =   54
      Text            =   "1"
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   54
      Left            =   4200
      TabIndex        =   53
      Text            =   "1"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   53
      Left            =   3720
      TabIndex        =   52
      Text            =   "1"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   52
      Left            =   3240
      TabIndex        =   51
      Text            =   "1"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   51
      Left            =   2640
      TabIndex        =   50
      Text            =   "1"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   50
      Left            =   2160
      TabIndex        =   49
      Text            =   "1"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   49
      Left            =   1680
      TabIndex        =   48
      Text            =   "1"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   48
      Left            =   1080
      TabIndex        =   47
      Text            =   "1"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   47
      Left            =   600
      TabIndex        =   46
      Text            =   "1"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   46
      Left            =   120
      TabIndex        =   45
      Text            =   "1"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   45
      Left            =   4200
      TabIndex        =   44
      Text            =   "1"
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   44
      Left            =   3720
      TabIndex        =   43
      Text            =   "1"
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   43
      Left            =   3240
      TabIndex        =   42
      Text            =   "1"
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   42
      Left            =   2640
      TabIndex        =   41
      Text            =   "1"
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   41
      Left            =   2160
      TabIndex        =   40
      Text            =   "1"
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   40
      Left            =   1680
      TabIndex        =   39
      Text            =   "1"
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   39
      Left            =   1080
      TabIndex        =   38
      Text            =   "1"
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   38
      Left            =   600
      TabIndex        =   37
      Text            =   "1"
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   37
      Left            =   120
      TabIndex        =   36
      Text            =   "1"
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   36
      Left            =   4200
      TabIndex        =   35
      Text            =   "1"
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   35
      Left            =   3720
      TabIndex        =   34
      Text            =   "1"
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   34
      Left            =   3240
      TabIndex        =   33
      Text            =   "1"
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   33
      Left            =   2640
      TabIndex        =   32
      Text            =   "1"
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   32
      Left            =   2160
      TabIndex        =   31
      Text            =   "1"
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   31
      Left            =   1680
      TabIndex        =   30
      Text            =   "1"
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   30
      Left            =   1080
      TabIndex        =   29
      Text            =   "1"
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   29
      Left            =   600
      TabIndex        =   28
      Text            =   "1"
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   28
      Left            =   120
      TabIndex        =   27
      Text            =   "1"
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   27
      Left            =   4200
      TabIndex        =   26
      Text            =   "1"
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   26
      Left            =   3720
      TabIndex        =   25
      Text            =   "1"
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   25
      Left            =   3240
      TabIndex        =   24
      Text            =   "1"
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   24
      Left            =   2640
      TabIndex        =   23
      Text            =   "1"
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   23
      Left            =   2160
      TabIndex        =   22
      Text            =   "1"
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   22
      Left            =   1680
      TabIndex        =   21
      Text            =   "1"
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   21
      Left            =   1080
      TabIndex        =   20
      Text            =   "1"
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   20
      Left            =   600
      TabIndex        =   19
      Text            =   "1"
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   19
      Left            =   120
      TabIndex        =   18
      Text            =   "1"
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   18
      Left            =   4200
      TabIndex        =   17
      Text            =   "1"
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   17
      Left            =   3720
      TabIndex        =   16
      Text            =   "1"
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   16
      Left            =   3240
      TabIndex        =   15
      Text            =   "1"
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   15
      Left            =   2640
      TabIndex        =   14
      Text            =   "1"
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   14
      Left            =   2160
      TabIndex        =   13
      Text            =   "1"
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   13
      Left            =   1680
      TabIndex        =   12
      Text            =   "1"
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   12
      Left            =   1080
      TabIndex        =   11
      Text            =   "1"
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   11
      Left            =   600
      TabIndex        =   10
      Text            =   "1"
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   10
      Left            =   120
      TabIndex        =   9
      Text            =   "1"
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   4200
      TabIndex        =   8
      Text            =   "1"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   3720
      TabIndex        =   7
      Text            =   "1"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   3240
      TabIndex        =   6
      Text            =   "1"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   2640
      TabIndex        =   5
      Text            =   "1"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   2160
      TabIndex        =   4
      Text            =   "1"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   1680
      TabIndex        =   3
      Text            =   "1"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   1080
      TabIndex        =   2
      Text            =   "1"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   600
      TabIndex        =   1
      Text            =   "1"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox TxtSudoku 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Text            =   "1"
      Top             =   840
      Width           =   495
   End
   Begin VB.Image ImB9 
      Appearance      =   0  '2D
      Height          =   495
      Left            =   4680
      MouseIcon       =   "FrmSudoku.frx":08CA
      MousePointer    =   99  'Benutzerdefiniert
      Top             =   4440
      Width           =   135
   End
   Begin VB.Image ImB8 
      Appearance      =   0  '2D
      Height          =   495
      Left            =   3120
      MouseIcon       =   "FrmSudoku.frx":0BD4
      MousePointer    =   99  'Benutzerdefiniert
      Top             =   4440
      Width           =   135
   End
   Begin VB.Image ImB7 
      Appearance      =   0  '2D
      Height          =   495
      Left            =   1560
      MouseIcon       =   "FrmSudoku.frx":0EDE
      MousePointer    =   99  'Benutzerdefiniert
      Top             =   4440
      Width           =   135
   End
   Begin VB.Image ImB6 
      Appearance      =   0  '2D
      Height          =   495
      Left            =   4680
      MouseIcon       =   "FrmSudoku.frx":11E8
      MousePointer    =   99  'Benutzerdefiniert
      Top             =   2880
      Width           =   135
   End
   Begin VB.Image ImB5 
      Appearance      =   0  '2D
      Height          =   495
      Left            =   3120
      MouseIcon       =   "FrmSudoku.frx":14F2
      MousePointer    =   99  'Benutzerdefiniert
      Top             =   2880
      Width           =   135
   End
   Begin VB.Image ImB4 
      Appearance      =   0  '2D
      Height          =   495
      Left            =   1560
      MouseIcon       =   "FrmSudoku.frx":17FC
      MousePointer    =   99  'Benutzerdefiniert
      Top             =   2880
      Width           =   135
   End
   Begin VB.Image ImC9 
      Appearance      =   0  '2D
      Height          =   135
      Left            =   4200
      MouseIcon       =   "FrmSudoku.frx":1B06
      MousePointer    =   99  'Benutzerdefiniert
      Top             =   720
      Width           =   495
   End
   Begin VB.Image ImC8 
      Appearance      =   0  '2D
      Height          =   135
      Left            =   3720
      MouseIcon       =   "FrmSudoku.frx":23D0
      MousePointer    =   99  'Benutzerdefiniert
      Top             =   720
      Width           =   495
   End
   Begin VB.Image ImC7 
      Appearance      =   0  '2D
      Height          =   135
      Left            =   3240
      MouseIcon       =   "FrmSudoku.frx":2C9A
      MousePointer    =   99  'Benutzerdefiniert
      Top             =   720
      Width           =   495
   End
   Begin VB.Image ImC6 
      Appearance      =   0  '2D
      Height          =   135
      Left            =   2640
      MouseIcon       =   "FrmSudoku.frx":3564
      MousePointer    =   99  'Benutzerdefiniert
      Top             =   720
      Width           =   495
   End
   Begin VB.Image ImC5 
      Appearance      =   0  '2D
      Height          =   135
      Left            =   2160
      MouseIcon       =   "FrmSudoku.frx":3E2E
      MousePointer    =   99  'Benutzerdefiniert
      Top             =   720
      Width           =   495
   End
   Begin VB.Image ImC4 
      Appearance      =   0  '2D
      Height          =   135
      Left            =   1680
      MouseIcon       =   "FrmSudoku.frx":46F8
      MousePointer    =   99  'Benutzerdefiniert
      Top             =   720
      Width           =   495
   End
   Begin VB.Image ImL9 
      Appearance      =   0  '2D
      Height          =   495
      Left            =   0
      MouseIcon       =   "FrmSudoku.frx":4FC2
      MousePointer    =   99  'Benutzerdefiniert
      Top             =   4920
      Width           =   135
   End
   Begin VB.Image ImL8 
      Appearance      =   0  '2D
      Height          =   495
      Left            =   0
      MouseIcon       =   "FrmSudoku.frx":52CC
      MousePointer    =   99  'Benutzerdefiniert
      Top             =   4440
      Width           =   135
   End
   Begin VB.Image ImL7 
      Appearance      =   0  '2D
      Height          =   495
      Left            =   0
      MouseIcon       =   "FrmSudoku.frx":55D6
      MousePointer    =   99  'Benutzerdefiniert
      Top             =   3960
      Width           =   135
   End
   Begin VB.Image ImL6 
      Appearance      =   0  '2D
      Height          =   495
      Left            =   0
      MouseIcon       =   "FrmSudoku.frx":58E0
      MousePointer    =   99  'Benutzerdefiniert
      Top             =   3360
      Width           =   135
   End
   Begin VB.Image ImL4 
      Appearance      =   0  '2D
      Height          =   495
      Left            =   0
      MouseIcon       =   "FrmSudoku.frx":5BEA
      MousePointer    =   99  'Benutzerdefiniert
      Top             =   2400
      Width           =   135
   End
   Begin VB.Image ImL5 
      Appearance      =   0  '2D
      Height          =   495
      Left            =   0
      MouseIcon       =   "FrmSudoku.frx":5EF4
      MousePointer    =   99  'Benutzerdefiniert
      Top             =   2880
      Width           =   135
   End
   Begin VB.Image ImB3 
      Appearance      =   0  '2D
      Height          =   495
      Left            =   4680
      MouseIcon       =   "FrmSudoku.frx":61FE
      MousePointer    =   99  'Benutzerdefiniert
      Top             =   1320
      Width           =   135
   End
   Begin VB.Image ImB2 
      Appearance      =   0  '2D
      Height          =   495
      Left            =   3120
      MouseIcon       =   "FrmSudoku.frx":6508
      MousePointer    =   99  'Benutzerdefiniert
      Top             =   1320
      Width           =   135
   End
   Begin VB.Image ImB1 
      Appearance      =   0  '2D
      Height          =   495
      Left            =   1560
      MouseIcon       =   "FrmSudoku.frx":6812
      MousePointer    =   99  'Benutzerdefiniert
      Top             =   1320
      Width           =   135
   End
   Begin VB.Image ImC3 
      Appearance      =   0  '2D
      Height          =   135
      Left            =   1080
      MouseIcon       =   "FrmSudoku.frx":6B1C
      MousePointer    =   99  'Benutzerdefiniert
      Top             =   720
      Width           =   495
   End
   Begin VB.Image ImC2 
      Appearance      =   0  '2D
      Height          =   135
      Left            =   600
      MouseIcon       =   "FrmSudoku.frx":73E6
      MousePointer    =   99  'Benutzerdefiniert
      Top             =   720
      Width           =   495
   End
   Begin VB.Image ImC1 
      Appearance      =   0  '2D
      Height          =   135
      Left            =   120
      MouseIcon       =   "FrmSudoku.frx":7CB0
      MousePointer    =   99  'Benutzerdefiniert
      Top             =   720
      Width           =   495
   End
   Begin VB.Image ImL3 
      Appearance      =   0  '2D
      Height          =   495
      Left            =   0
      MouseIcon       =   "FrmSudoku.frx":857A
      MousePointer    =   99  'Benutzerdefiniert
      Top             =   1800
      Width           =   135
   End
   Begin VB.Image ImL2 
      Appearance      =   0  '2D
      Height          =   495
      Left            =   0
      MouseIcon       =   "FrmSudoku.frx":8884
      MousePointer    =   99  'Benutzerdefiniert
      Top             =   1320
      Width           =   135
   End
   Begin VB.Image ImL1 
      Appearance      =   0  '2D
      Height          =   495
      Left            =   0
      MouseIcon       =   "FrmSudoku.frx":8B8E
      MousePointer    =   99  'Benutzerdefiniert
      Top             =   840
      Width           =   135
   End
   Begin VB.Label LblSudoku 
      Alignment       =   1  'Rechts
      Caption         =   "0 / 81"
      Height          =   255
      Left            =   4080
      TabIndex        =   88
      Top             =   480
      Width           =   615
   End
End
Attribute VB_Name = "FrmSudoku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '733 Zeilen 'halt nur 397 Zeilen, jetzt in Res mehr
Private mSudoku As SudokuGame
Private mUndoRedo As SGUndoRedo
Private mn As Long
Private mRandomGames As New RandomGames

Private Sub BtnInfo_Click()
  frmAbout.Show 1, Me
End Sub

Private Sub RndExample_Click()
Dim StrG As String
  StrG = mRandomGames.GetNextGame
  mSudoku.GameName = mRandomGames.GameName
  mSudoku.ReadNParseFromStr (StrG)
  mSudoku.Show
  SetLevel (mSudoku.GameLevel)
  SetGameTitle (mSudoku.GameName)
End Sub
Private Sub SetGameTitle(Name As String)
Dim StrCap As String
  StrCap = "SuMBOku 1.2"
  If Len(Name) > 0 Then
    StrCap = StrCap & " - " & Name
  End If
  Me.Caption = StrCap
End Sub
'prozedure zum zusammenfügen der sudoku-Dateien zur StringRessource
'Private Sub Command1_Click()
'Dim OFD As New OpenFileDialog
'Dim P As String
'Dim PFNAll As String
'Dim PFNCurSS 'As String
'Dim FNr As Integer
'Dim InpStr As String, LInStr
'Dim AllStr As String
'Dim FNam As String
'TryE: On Error GoTo Catch
'  P = App.Path & "\ss-files\"
'  PFNAll = "AllSudoku_ss.txt"
'  OFD.MultiSelect = True
'  OFD.InitialDirectory = P
'  If OFD.ShowDialog = DialogResult_OK Then
'    For Each PFNCurSS In OFD.FileNames
'      FNr = FreeFile
'      Open PFNCurSS For Input As #FNr
'      InpStr = vbNullString
'      Do While Not EOF(FNr)
'        Line Input #FNr, LInStr
'        InpStr = InpStr & LInStr
'      Loop
'      Close #FNr
'      FNam = Right$(PFNCurSS, Len(PFNCurSS) - InStrRev(PFNCurSS, "\")) & ", "
'      AllStr = AllStr & FNam & InpStr & vbCrLf
'    Next
'    FNr = FreeFile
'    Open P & "\" & PFNAll For Binary Access Write As #FNr
'    Put #FNr, , AllStr
'    Close #FNr
'  End If
'  Exit Sub
'Catch:
'  Close #FNr
'End Sub

'#################'   I n i t   T h e   G a m e   '#################'
Private Sub Form_Load()
  'App.HelpFile = App.Path & "\" & "SuMBOku.hlp"
  Set mSudoku = New SudokuGame
  Set mUndoRedo = New SGUndoRedo
  mn = 3
  Call mSudoku.NewC(mn * mn)
  Set mUndoRedo.Sudoku = mSudoku
  Call mUndoRedo.SetBtnUndoRedo(BtnUndo, BtnRedo)
  'BtnUndo.Enabled = False
  'BtnRedo.Enabled = False
  'Set ImL4.Picture = LoadResPicture(202, vbResBitmap)
  PrepareImgs
End Sub
Private Sub PrepareImgs()
Dim Img As Image
Dim i As Long
  For i = 1 To mn * mn
    Set Img = Me.Controls("ImB" & CStr(i))
    Set Img.Picture = LoadResPicture(201, vbResBitmap)
    Set Img = Me.Controls("ImL" & CStr(i))
    Set Img.Picture = LoadResPicture(202, vbResBitmap)
    Set Img = Me.Controls("ImC" & CStr(i))
    Set Img.Picture = LoadResPicture(203, vbResBitmap)
  Next
End Sub
Public Sub SetupGame()
  Call mSudoku.SetTxtBoxes(TxtSudoku)
  Call mSudoku.Clear
  Call SetLevel(0)
  SetImgs
End Sub
Private Sub SetImgs()
Dim BLC As SudokuBLC
Dim i As Long
  For i = 1 To mn * mn
    Set BLC = mSudoku.mBlockCol(i): Set BLC.Img = Me.Controls("ImB" & CStr(i))
  Next
  For i = 1 To mn * mn
    Set BLC = mSudoku.mLineCol(i): Set BLC.Img = Me.Controls("ImL" & CStr(i))
  Next
  For i = 1 To mn * mn
    Set BLC = mSudoku.mColmCol(i): Set BLC.Img = Me.Controls("ImC" & CStr(i))
  Next
End Sub
'#################'        Commands  Oben         '#################'
Private Sub Image1_Click(Index As Integer)
 Select Case Index
  Case 0, 4: Call SetLevel(1)
  Case 1, 5: Call SetLevel(2)
  Case 2, 6: Call SetLevel(3)
  Case 3, 7: Call SetLevel(4)
  Case 8: Call SetLevel(0)
  End Select
End Sub
Private Sub Image2_Click()

End Sub

Private Sub Image3_Click()

End Sub

Private Sub Image4_Click()

End Sub

Private Sub Image5_Click()

End Sub


Private Sub ImGL0_Click(): Call SetLevel(0): End Sub
Private Sub ImGL1_Click(): Call SetLevel(1): End Sub
Private Sub ImGL2_Click(): Call SetLevel(2): End Sub
Private Sub ImGL3_Click(): Call SetLevel(3): End Sub
Private Sub ImGL4_Click(): Call SetLevel(4): End Sub
Private Sub SetLevel(Index As Long)
Dim i As Long, n As Long
Dim Im As Image
  mSudoku.GameLevel = Index
  n = 4
  For i = 1 To Index
    Set Im = Me.Controls("ImGL" & CStr(i))
    Set Im.Picture = LoadResPicture(12, vbResBitmap)
    'im.Picture.hPal =
  Next
  For i = Index + 1 To n
    Set Im = Me.Controls("ImGL" & CStr(i))
    Set Im.Picture = LoadResPicture(11, vbResBitmap)
  Next
End Sub

Private Sub BtnOptions_Click()
  Load frmSudokuOptions
  Set frmSudokuOptions.Sudoku = mSudoku
  frmSudokuOptions.Show 1, Me
End Sub
Private Sub ImB1_Click(): Call BLCVDblClickHandlerShowFrmSuEdit("B", 1, ImB1.Left, ImB1.Top): End Sub
Private Sub ImB2_Click(): Call BLCVDblClickHandlerShowFrmSuEdit("B", 2, ImB2.Left, ImB2.Top): End Sub
Private Sub ImB3_Click(): Call BLCVDblClickHandlerShowFrmSuEdit("B", 3, ImB3.Left, ImB3.Top): End Sub
Private Sub ImB4_Click(): Call BLCVDblClickHandlerShowFrmSuEdit("B", 4, ImB4.Left, ImB4.Top): End Sub
Private Sub ImB5_Click(): Call BLCVDblClickHandlerShowFrmSuEdit("B", 5, ImB5.Left, ImB5.Top): End Sub
Private Sub ImB6_Click(): Call BLCVDblClickHandlerShowFrmSuEdit("B", 6, ImB6.Left, ImB6.Top): End Sub
Private Sub ImB7_Click(): Call BLCVDblClickHandlerShowFrmSuEdit("B", 7, ImB7.Left, ImB7.Top): End Sub
Private Sub ImB8_Click(): Call BLCVDblClickHandlerShowFrmSuEdit("B", 8, ImB8.Left, ImB8.Top): End Sub
Private Sub ImB9_Click(): Call BLCVDblClickHandlerShowFrmSuEdit("B", 9, ImB9.Left, ImB9.Top): End Sub

Private Sub ImL1_Click(): Call BLCVDblClickHandlerShowFrmSuEdit("L", 1, ImL1.Left, ImL1.Top): End Sub
Private Sub ImL2_Click(): Call BLCVDblClickHandlerShowFrmSuEdit("L", 2, ImL2.Left, ImL2.Top): End Sub
Private Sub ImL3_Click(): Call BLCVDblClickHandlerShowFrmSuEdit("L", 3, ImL3.Left, ImL3.Top): End Sub
Private Sub ImL4_Click(): Call BLCVDblClickHandlerShowFrmSuEdit("L", 4, ImL4.Left, ImL4.Top): End Sub
Private Sub ImL5_Click(): Call BLCVDblClickHandlerShowFrmSuEdit("L", 5, ImL5.Left, ImL5.Top): End Sub
Private Sub ImL6_Click(): Call BLCVDblClickHandlerShowFrmSuEdit("L", 6, ImL6.Left, ImL6.Top): End Sub
Private Sub ImL7_Click(): Call BLCVDblClickHandlerShowFrmSuEdit("L", 7, ImL7.Left, ImL7.Top): End Sub
Private Sub ImL8_Click(): Call BLCVDblClickHandlerShowFrmSuEdit("L", 8, ImL8.Left, ImL8.Top): End Sub
Private Sub ImL9_Click(): Call BLCVDblClickHandlerShowFrmSuEdit("L", 9, ImL9.Left, ImL9.Top): End Sub

Private Sub ImC1_Click(): Call BLCVDblClickHandlerShowFrmSuEdit("C", 1, ImC1.Left, ImC1.Top): End Sub
Private Sub ImC2_Click(): Call BLCVDblClickHandlerShowFrmSuEdit("C", 2, ImC2.Left, ImC2.Top): End Sub
Private Sub ImC3_Click(): Call BLCVDblClickHandlerShowFrmSuEdit("C", 3, ImC3.Left, ImC3.Top): End Sub
Private Sub ImC4_Click(): Call BLCVDblClickHandlerShowFrmSuEdit("C", 4, ImC4.Left, ImC4.Top): End Sub
Private Sub ImC5_Click(): Call BLCVDblClickHandlerShowFrmSuEdit("C", 5, ImC5.Left, ImC5.Top): End Sub
Private Sub ImC6_Click(): Call BLCVDblClickHandlerShowFrmSuEdit("C", 6, ImC6.Left, ImC6.Top): End Sub
Private Sub ImC7_Click(): Call BLCVDblClickHandlerShowFrmSuEdit("C", 7, ImC7.Left, ImC7.Top): End Sub
Private Sub ImC8_Click(): Call BLCVDblClickHandlerShowFrmSuEdit("C", 8, ImC8.Left, ImC8.Top): End Sub
Private Sub ImC9_Click(): Call BLCVDblClickHandlerShowFrmSuEdit("C", 9, ImC9.Left, ImC9.Top): End Sub

Private Sub TxtSudoku_DblClick(Index As Integer)
  Call BLCVDblClickHandlerShowFrmSuEdit("V", CLng(Index))
End Sub

Private Sub BLCVDblClickHandlerShowFrmSuEdit(Typ As String, Index As Long, Optional L As Long, Optional T As Long)
Dim StrCap As String
Dim mBLCV As SudokuMissPoss 'SudokuBLC
Dim CurCtrl As Control
Dim FL As Long, FT As Long
  Select Case Asc(Typ)
  Case 66 '"B"
    StrCap = "Block "
    Set mBLCV = mSudoku.mBlockCol(Index)
  Case 67 '"C"
    StrCap = "Column "
    Set mBLCV = mSudoku.mColmCol(Index)
  Case 76 '"L"
    StrCap = "Line "
    Set mBLCV = mSudoku.mLineCol(Index)
  Case 86 '"V"
    StrCap = "Cell "
    Set mBLCV = mSudoku.ValueCol(Index)
  End Select
  frmSuEdit.Caption = StrCap & CStr(Index) & ":"
  Set frmSuEdit.BLCV = mBLCV
  If L = 0 And T = 0 Then
    Set CurCtrl = Me.ActiveControl
    L = CurCtrl.Left '/ Screen.TwipsPerPixelX
    T = CurCtrl.Top '/ Screen.TwipsPerPixelY
  End If
  FL = Me.Left + L * Screen.TwipsPerPixelX
  FT = Me.Top + T * Screen.TwipsPerPixelY - frmSuEdit.Height / 2
  frmSuEdit.Move FL, FT
  frmSuEdit.Show 1, Me
  
End Sub


Private Sub LblSudoku_Click()
  Call UpdateLabelSudoku
End Sub
Private Sub UpdateLabelSudoku()
Dim bb As Long
  bb = mSudoku.GetAmountOfUnZeroVals
  LblSudoku.Caption = CStr(bb) & " / " & CStr(81)
End Sub


'#################'        Commands  Unten         '#################'
Private Sub BtnClear_Click()
  Call mUndoRedo.SaveCompleteUndo
  mSudoku.Clear
  mSudoku.Show
  SetLevel (0)
End Sub

Private Sub BtnUndo_Click()
  mUndoRedo.UndoLastAction
  mSudoku.Show
End Sub
Private Sub BtnRedo_Click()
  mUndoRedo.RedoLastAction
  mSudoku.Show
End Sub

Private Sub BtnSolve_Click()
  'if msudoku.GetAmountOfUnZeroVals <
  mSudoku.Solve 'Rätsel lösen
  mSudoku.Show  'Ergebnisse anzeigen
  
  UpdateLabelSudoku 'Anzahl der gelösten Zellen
  'zum Schluß prüfen, ob das Rätsel richtig und ob schon fertig,
  'nächste Iteration soll der User selber machen, durch nochmaliges
  'Klicken auf Solve
  mSudoku.DoTheCheck
End Sub

Private Sub BtnOpen_Click()
Dim OFD As New OpenFileDialog
  OFD.Filter = "SuMBOku-files [*.smbk]|*.smbk|simplesudoku [*.ss]|*.ss"
  OFD.InitialDirectory = App.Path
  If OFD.ShowDialog(Me) = VbMsgBoxResult.vbOK Then
    Call mSudoku.ReadFromFile(OFD.FileName)
    mSudoku.Show
    SetLevel (mSudoku.GameLevel)
    SetGameTitle (mSudoku.GameName)
  End If
End Sub
Private Sub BtnSave_Click()
Dim SFD As New SaveFileDialog
  SFD.DefaultExt = ".smbk"
  SFD.Filter = "SuMBOku-files [*.smbk]|*.smbk"
  SFD.InitialDirectory = App.Path
  If SFD.ShowDialog(Me) = VbMsgBoxResult.vbOK Then
    Call mSudoku.WriteToFile(SFD.FileName)
  End If
End Sub

Private Sub BtnExit_Click()
  Unload Me
End Sub

'Private Sub BtnTestLines_Click()
'  mSudoku.TestLines
'End Sub
'Private Sub BtnTestColms_Click()
'  mSudoku.TestColms
'End Sub
'Private Sub BtnTestBlocks_Click()
'  mSudoku.TestBlocks
'End Sub

Private Sub TxtSudoku_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim i As Long, n2 As Long, n4 As Long, d As Long
Dim sh As Long, se As Long, su As Long, sd As Long
  n2 = mn * mn
  n4 = n2 * n2
  'If Len(TxtSudoku(Index).Text) > 0 Then
  'KeyCode = vbKeyRight
  'If Shift And ctrl Then
  i = Index
  d = Index Mod n2
  If d Then sh = 1 Else sh = -n2 + 1
  If d Then se = n2 Else se = 0
  Select Case KeyCode
  Case vbKeyHome:     i = i - d + sh 'zur ersten  Zelle in der Zeile  springen
  Case vbKeyEnd:      i = i - d + se 'zur letzten Zelle in der Zeile  springen
  Case vbKeyPageUp:   i = d - sh + 1 'zur ersten  Zelle in der Spalte springen
  Case vbKeyPageDown: i = n4 - n2 + d - sh + 1 'zur letzten Zelle in der Spalte springen
  Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyReturn
    Select Case KeyCode
    Case vbKeyUp
      i = Index - n2 '9
      If i < 1 Then i = n4 + i '81
    Case vbKeyDown, vbKeyReturn
      i = Index + n2 '9
      If i > n4 Then i = i - n4 '81
    Case vbKeyLeft
      i = Index - 1
      If i < 1 Then i = n4 '81
    Case vbKeyRight
      i = Index + 1
      If i > n4 Then i = 1 '81
    End Select
  Case vbKeyDelete
  Case vbKeyBack
    'MsgBox "vbKeyBack"
  Case Else
    Select Case KeyCode
    Case vbKeyF1, vbKeyNumpad1: KeyCode = 49
    Case vbKeyF2, vbKeyNumpad2: KeyCode = 50
    Case vbKeyF3, vbKeyNumpad3: KeyCode = 51
    Case vbKeyF4, vbKeyNumpad4: KeyCode = 52
    Case vbKeyF5, vbKeyNumpad5: KeyCode = 53
    Case vbKeyF6, vbKeyNumpad6: KeyCode = 54
    Case vbKeyF7, vbKeyNumpad7: KeyCode = 55
    Case vbKeyF8, vbKeyNumpad8: KeyCode = 56
    Case vbKeyF9, vbKeyNumpad9: KeyCode = 57
    Case vbKeyF10, vbKeyNumpad0: KeyCode = 58
    Case vbKeyA, vbKeyB, vbKeyC, vbKeyD, vbKeyE, vbKeyF, vbKeyG
    Case Else: KeyCode = 0
    End Select
  End Select
  If i > n2 * n2 Then i = 1 'n2 * n2
  If i < 1 Then i = n2 * n2
  TxtSudoku(i).SetFocus
  'KeyCode(vbkey1) = 49
  'KeyCode(vbKeyF1) = 112
  'keycode(vbKeyNumpad1) = 97
End Sub

Private Sub TxtSudoku_KeyPress(Index As Integer, KeyAscii As Integer)
  Select Case UCase(Chr(KeyAscii))
  Case "1", "2", "3", "4", "5", "6", "7", "8", "9"
  Case "J": KeyAscii = Asc("1")
  Case "K": KeyAscii = Asc("2")
  Case "L": KeyAscii = Asc("3")
  Case "U": KeyAscii = Asc("4")
  Case "I": KeyAscii = Asc("5")
  Case "O": KeyAscii = Asc("6")
  Case "A", "B", "C", "D", "E", "F", "G"
  Case Else
    'If KeyAscii = vbKeyF1 Then MsgBox "F1"
    If Not KeyAscii = vbKeyBack Then KeyAscii = 0 'Asc(0)
  End Select
End Sub
Private Sub TxtSudoku_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim NumVal As Double
Dim V As SudokuVal
  If Not KeyCode = vbKeyBack Then
    If IsNumeric(TxtSudoku(Index).Text) Then
      NumVal = CDbl(TxtSudoku(Index).Text)
      If (0 < NumVal) And (NumVal < 10) Then
        Set V = mSudoku.ValueCol(Index)
        V.Value = CLng(NumVal)
      End If
    End If
  End If
  If Len(TxtSudoku(Index).Text) = 0 Then
    Set V = mSudoku.ValueCol(Index)
    V.Value = CLng(0)
  End If
End Sub

'Private Sub BtnExample1_Click()
'Dim V As SudokuVal
'Dim i As Long
'  Call mUndoRedo.SaveCompleteUndo
'  Call mSudoku.Clear
'  For i = 1 To 81
'    Set V = mSudoku.ValueCol(i)
'    If i = 1 Then V.Value = 6
'    'If i = 2 Then V.Value = 0
'    'If i = 3 Then V.Value = 0
'    If i = 4 Then V.Value = 8
'    If i = 5 Then V.Value = 3
'    If i = 6 Then V.Value = 4
'    If i = 7 Then V.Value = 1
'    'If i = 8 Then V.Value = 0
'    'If i = 9 Then V.Value = 0
'    'If i = 10 Then V.Value = 0
'    If i = 11 Then V.Value = 9
'    If i = 12 Then V.Value = 1
'    'If i = 13 Then V.Value = 0
'    'If i = 14 Then V.Value = 0
'    If i = 15 Then V.Value = 7
'    If i = 16 Then V.Value = 8
'    'If i = 17 Then V.Value = 0
'    If i = 18 Then V.Value = 4
'    'If i = 19 Then V.Value = 0
'    'If i = 20 Then V.Value = 0
'    'If i = 21 Then V.Value = 0
'    'If i = 22 Then V.Value = 0
'    If i = 23 Then V.Value = 9
'    'If i = 24 Then V.Value = 0
'    'If i = 25 Then V.Value = 0
'    If i = 26 Then V.Value = 6
'    'If i = 27 Then V.Value = 0
'    'If i = 28 Then V.Value = 0
'    'If i = 29 Then V.Value = 0
'    If i = 30 Then V.Value = 8
'    If i = 31 Then V.Value = 0
'    'If i = 32 Then V.Value = 0
'    If i = 33 Then V.Value = 3
'    'If i = 34 Then V.Value = 0
'    If i = 35 Then V.Value = 4
'    If i = 36 Then V.Value = 9
'    If i = 37 Then V.Value = 2
'    'If i = 38 Then V.Value = 0
'    'If i = 39 Then V.Value = 0
'    'If i = 40 Then V.Value = 0
'    'If i = 41 Then V.Value = 0
'    'If i = 42 Then V.Value = 0
'    'If i = 43 Then V.Value = 0
'    'If i = 44 Then V.Value = 0
'    If i = 45 Then V.Value = 3
'    If i = 46 Then V.Value = 3
'    If i = 47 Then V.Value = 1
'    'If i = 48 Then V.Value = 0
'    If i = 49 Then V.Value = 5
'    'If i = 50 Then V.Value = 0
'    If i = 51 Then V.Value = 6
'    If i = 52 Then V.Value = 7
'    'If i = 53 Then V.Value = 0
'    'If i = 54 Then V.Value = 0
'    'If i = 55 Then V.Value = 0
'    If i = 56 Then V.Value = 8
'    'If i = 57 Then V.Value = 0
'    'If i = 58 Then V.Value = 0
'    If i = 59 Then V.Value = 5
'    'If i = 60 Then V.Value = 0
'    'If i = 61 Then V.Value = 0
'    'If i = 62 Then V.Value = 0
'    'If i = 63 Then V.Value = 0
'    If i = 64 Then V.Value = 9
'    'If i = 65 Then V.Value = 0
'    If i = 66 Then V.Value = 5
'    If i = 67 Then V.Value = 3
'    'If i = 68 Then V.Value = 0
'    'If i = 69 Then V.Value = 0
'    If i = 70 Then V.Value = 4
'    If i = 71 Then V.Value = 2
'    'If i = 72 Then V.Value = 0
'    'If i = 73 Then V.Value = 0
'    'If i = 74 Then V.Value = 0
'    If i = 75 Then V.Value = 2
'    If i = 76 Then V.Value = 9
'    'If i = 77 Then V.Value = 0
'    If i = 78 Then V.Value = 8
'    'If i = 79 Then V.Value = 0
'    'If i = 80 Then V.Value = 0
'    If i = 81 Then V.Value = 1
'  Next
'  mSudoku.Show
'  Call SetLevel(1)
'  UpdateLabelSudoku
'End Sub
'
'Private Sub BtnExample2_Click()
'Dim V As SudokuVal
'Dim i As Long
'  Call mUndoRedo.SaveCompleteUndo
'  Call mSudoku.Clear
'  For i = 1 To 81
'    Set V = mSudoku.ValueCol(i)
'    If i = 1 Then V.Value = 1
''    If i = 2 Then V.Value = 0
''    If i = 3 Then V.Value = 0
'    If i = 4 Then V.Value = 4
''    If i = 5 Then V.Value = 0
'    If i = 6 Then V.Value = 5
''    If i = 7 Then V.Value = 0
''    If i = 8 Then V.Value = 0
'    If i = 9 Then V.Value = 9
''    If i = 10 Then V.Value = 0
''    If i = 11 Then V.Value = 0
'    If i = 12 Then V.Value = 3
'    If i = 13 Then V.Value = 1
''    If i = 14 Then V.Value = 0
'    If i = 15 Then V.Value = 7
'    If i = 16 Then V.Value = 5
''    If i = 17 Then V.Value = 0
''    If i = 18 Then V.Value = 0
''    If i = 19 Then V.Value = 0
''    If i = 20 Then V.Value = 0
'    If i = 21 Then V.Value = 9
''    If i = 22 Then V.Value = 0
''    If i = 23 Then V.Value = 0
''    If i = 24 Then V.Value = 0
'    If i = 25 Then V.Value = 4
''    If i = 26 Then V.Value = 0
''    If i = 27 Then V.Value = 0
'    If i = 28 Then V.Value = 3
''    If i = 29 Then V.Value = 0
''    If i = 30 Then V.Value = 0
'    If i = 31 Then V.Value = 6
'    If i = 32 Then V.Value = 2
'    If i = 33 Then V.Value = 8
''    If i = 34 Then V.Value = 0
''    If i = 35 Then V.Value = 0
'    If i = 36 Then V.Value = 7
''    If i = 37 Then V.Value = 0
'    If i = 38 Then V.Value = 4
''    If i = 39 Then V.Value = 0
''    If i = 40 Then V.Value = 0
''    If i = 41 Then V.Value = 0
''    If i = 42 Then V.Value = 0
''    If i = 43 Then V.Value = 0
'    If i = 44 Then V.Value = 5
''    If i = 45 Then V.Value = 0
'    If i = 46 Then V.Value = 6
''    If i = 47 Then V.Value = 0
''    If i = 48 Then V.Value = 0
'    If i = 49 Then V.Value = 5
'    If i = 50 Then V.Value = 4
'    If i = 51 Then V.Value = 3
''    If i = 52 Then V.Value = 0
''    If i = 53 Then V.Value = 0
'    If i = 54 Then V.Value = 8
''    If i = 55 Then V.Value = 0
''    If i = 56 Then V.Value = 0
'    If i = 57 Then V.Value = 5
''    If i = 58 Then V.Value = 0
''    If i = 59 Then V.Value = 0
''    If i = 60 Then V.Value = 0
'    If i = 61 Then V.Value = 1
''    If i = 62 Then V.Value = 0
''    If i = 63 Then V.Value = 0
''    If i = 64 Then V.Value = 0
''    If i = 65 Then V.Value = 0
'    If i = 66 Then V.Value = 4
'    If i = 67 Then V.Value = 2
''    If i = 68 Then V.Value = 0
'    If i = 69 Then V.Value = 1
'    If i = 70 Then V.Value = 6
''    If i = 71 Then V.Value = 0
''    If i = 72 Then V.Value = 0
'    If i = 73 Then V.Value = 2
''    If i = 74 Then V.Value = 0
''    If i = 75 Then V.Value = 0
'    If i = 76 Then V.Value = 9
''    If i = 77 Then V.Value = 0
'    If i = 78 Then V.Value = 4
''    If i = 79 Then V.Value = 0
''    If i = 80 Then V.Value = 0
'    If i = 81 Then V.Value = 5
'  Next
'  mSudoku.Show
'  Call SetLevel(2)
'  UpdateLabelSudoku
'End Sub
'
'Private Sub BtnExample3_Click()
'Dim V As SudokuVal
'Dim i As Long
'  Call mUndoRedo.SaveCompleteUndo
'  Call mSudoku.Clear
'  For i = 1 To 81
'    Set V = mSudoku.ValueCol(i)
''    If i = 1 Then V.Value = 0
'    If i = 2 Then V.Value = 7
''    If i = 3 Then V.Value = 0
'    If i = 4 Then V.Value = 4
''    If i = 5 Then V.Value = 0
'    If i = 6 Then V.Value = 9
'    If i = 7 Then V.Value = 1
''    If i = 8 Then V.Value = 0
''    If i = 9 Then V.Value = 0
''    If i = 10 Then V.Value = 0
'    If i = 11 Then V.Value = 1
''    If i = 12 Then V.Value = 0
''    If i = 13 Then V.Value = 0
''    If i = 14 Then V.Value = 0
'    If i = 15 Then V.Value = 2
''    If i = 16 Then V.Value = 0
'    If i = 17 Then V.Value = 9
'    If i = 18 Then V.Value = 4
'    If i = 19 Then V.Value = 8
''    If i = 20 Then V.Value = 0
''    If i = 21 Then V.Value = 0
''    If i = 22 Then V.Value = 0
'    If i = 23 Then V.Value = 1
''    If i = 24 Then V.Value = 0
''    If i = 25 Then V.Value = 0
''    If i = 26 Then V.Value = 0
''    If i = 27 Then V.Value = 0
'    If i = 28 Then V.Value = 7
'    If i = 29 Then V.Value = 4
''    If i = 30 Then V.Value = 0
''    If i = 31 Then V.Value = 0
''    If i = 32 Then V.Value = 0
''    If i = 33 Then V.Value = 0
''    If i = 34 Then V.Value = 0
''    If i = 35 Then V.Value = 0
'    If i = 36 Then V.Value = 2
''    If i = 37 Then V.Value = 0
''    If i = 38 Then V.Value = 0
'    If i = 39 Then V.Value = 6
''    If i = 40 Then V.Value = 0
'    If i = 41 Then V.Value = 5
''    If i = 42 Then V.Value = 0
'    If i = 43 Then V.Value = 3
''    If i = 44 Then V.Value = 0
''    If i = 45 Then V.Value = 0
'    If i = 46 Then V.Value = 1
''    If i = 47 Then V.Value = 0
''    If i = 48 Then V.Value = 0
''    If i = 49 Then V.Value = 0
''    If i = 50 Then V.Value = 0
''    If i = 51 Then V.Value = 0
''    If i = 52 Then V.Value = 0
'    If i = 53 Then V.Value = 4
'    If i = 54 Then V.Value = 8
''    If i = 55 Then V.Value = 0
''    If i = 56 Then V.Value = 0
''    If i = 57 Then V.Value = 0
''    If i = 58 Then V.Value = 0
'    If i = 59 Then V.Value = 2
''    If i = 60 Then V.Value = 0
''    If i = 61 Then V.Value = 0
''    If i = 62 Then V.Value = 0
'    If i = 63 Then V.Value = 3
'    If i = 64 Then V.Value = 3
'    If i = 65 Then V.Value = 2
''    If i = 66 Then V.Value = 0
'    If i = 67 Then V.Value = 7
''    If i = 68 Then V.Value = 0
''    If i = 69 Then V.Value = 0
''    If i = 70 Then V.Value = 0
'    If i = 71 Then V.Value = 6
''    If i = 72 Then V.Value = 0
''    If i = 73 Then V.Value = 0
''    If i = 74 Then V.Value = 0
'    If i = 75 Then V.Value = 8
'    If i = 76 Then V.Value = 6
''    If i = 77 Then V.Value = 0
'    If i = 78 Then V.Value = 3
''    If i = 79 Then V.Value = 0
'    If i = 80 Then V.Value = 7
''    If i = 81 Then V.Value = 0
'  Next
'  mSudoku.Show
'  Call SetLevel(3)
'  UpdateLabelSudoku
'End Sub
'
'Private Sub BtnExample4_Click()
'Dim V As SudokuVal
'Dim i As Long
'  Call mUndoRedo.SaveCompleteUndo
'  Call mSudoku.Clear
'  For i = 1 To 81
'    Set V = mSudoku.ValueCol(i)
''    If i = 1 Then V.Value = 0
''    If i = 2 Then V.Value = 0
'    If i = 3 Then V.Value = 4
''    If i = 4 Then V.Value = 0
''    If i = 5 Then V.Value = 0
''    If i = 6 Then V.Value = 0
''    If i = 7 Then V.Value = 0
'    If i = 8 Then V.Value = 3
''    If i = 9 Then V.Value = 0
'    If i = 10 Then V.Value = 1
''    If i = 11 Then V.Value = 0
''    If i = 12 Then V.Value = 0
''    If i = 13 Then V.Value = 0
'    If i = 14 Then V.Value = 5
'    If i = 15 Then V.Value = 3
'    If i = 16 Then V.Value = 2
''    If i = 17 Then V.Value = 0
'    If i = 18 Then V.Value = 6
''    If i = 19 Then V.Value = 0
'    If i = 20 Then V.Value = 6
''    If i = 21 Then V.Value = 0
''    If i = 22 Then V.Value = 0
''    If i = 23 Then V.Value = 0
'    If i = 24 Then V.Value = 1
''    If i = 25 Then V.Value = 0
''    If i = 26 Then V.Value = 0
''    If i = 27 Then V.Value = 0
''    If i = 28 Then V.Value = 0
''    If i = 29 Then V.Value = 0
''    If i = 30 Then V.Value = 0
''    If i = 31 Then V.Value = 0
'    If i = 32 Then V.Value = 2
'    If i = 33 Then V.Value = 5
''    If i = 34 Then V.Value = 0
'    If i = 35 Then V.Value = 4
''    If i = 36 Then V.Value = 0
''    If i = 37 Then V.Value = 0
'    If i = 38 Then V.Value = 2
''    If i = 39 Then V.Value = 0
''    If i = 40 Then V.Value = 0
''    If i = 41 Then V.Value = 0
''    If i = 42 Then V.Value = 0
''    If i = 43 Then V.Value = 0
'    If i = 44 Then V.Value = 6
''    If i = 45 Then V.Value = 0
''    If i = 46 Then V.Value = 0
'    If i = 47 Then V.Value = 1
''    If i = 48 Then V.Value = 0
'    If i = 49 Then V.Value = 8
'    If i = 50 Then V.Value = 7
''    If i = 51 Then V.Value = 0
''    If i = 52 Then V.Value = 0
''    If i = 53 Then V.Value = 0
''    If i = 54 Then V.Value = 0
''    If i = 55 Then V.Value = 0
''    If i = 56 Then V.Value = 0
''    If i = 57 Then V.Value = 0
'    If i = 58 Then V.Value = 5
''    If i = 59 Then V.Value = 0
''    If i = 60 Then V.Value = 0
''    If i = 61 Then V.Value = 0
'    If i = 62 Then V.Value = 7
''    If i = 63 Then V.Value = 0
'    If i = 64 Then V.Value = 8
''    If i = 65 Then V.Value = 0
'    If i = 66 Then V.Value = 6
'    If i = 67 Then V.Value = 7
'    If i = 68 Then V.Value = 3
''    If i = 69 Then V.Value = 0
''    If i = 70 Then V.Value = 0
''    If i = 71 Then V.Value = 0
'    If i = 72 Then V.Value = 5
''    If i = 73 Then V.Value = 0
'    If i = 74 Then V.Value = 5
''    If i = 75 Then V.Value = 0
''    If i = 76 Then V.Value = 0
''    If i = 77 Then V.Value = 0
''    If i = 78 Then V.Value = 0
'    If i = 79 Then V.Value = 9
''    If i = 80 Then V.Value = 0
''    If i = 81 Then V.Value = 0
'  Next
'  mSudoku.Show
'  Call SetLevel(4)
'  UpdateLabelSudoku
'End Sub
