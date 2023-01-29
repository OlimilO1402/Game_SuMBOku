VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Info"
   ClientHeight    =   3660
   ClientLeft      =   2340
   ClientTop       =   1815
   ClientWidth     =   4815
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   244
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   321
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PbMBOIngcom 
      BorderStyle     =   0  'Kein
      Height          =   255
      Left            =   2280
      ScaleHeight     =   255
      ScaleWidth      =   975
      TabIndex        =   4
      Top             =   3240
      Width           =   975
      Begin VB.Label LblMBOINGCOM 
         AutoSize        =   -1  'True
         Caption         =   "MBO-Ing.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   0
         TabIndex        =   5
         ToolTipText     =   "http://www.mbo-ing.com"
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3480
      TabIndex        =   0
      Top             =   3240
      Width           =   1260
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Oben ausrichten
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   321
      TabIndex        =   3
      Top             =   0
      Width           =   4815
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   3720
         Top             =   120
      End
      Begin VB.Image ImgSunRun 
         Height          =   675
         Left            =   120
         Picture         =   "frmAbout.frx":0000
         Top             =   120
         Width           =   675
      End
      Begin VB.Image ImgSuMBOku 
         Height          =   600
         Left            =   960
         Picture         =   "frmAbout.frx":0573
         Top             =   240
         Width           =   2010
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   0
         X2              =   369.933
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderStyle     =   6  'Innen ausgefüllt
         Index           =   1
         X1              =   0
         X2              =   371
         Y1              =   64
         Y2              =   64
      End
   End
   Begin VB.Label lblDescription 
      Caption         =   "Trainprogram Nr. 33"
      ForeColor       =   &H00000000&
      Height          =   1890
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   4605
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   225
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   1965
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '40 Zeilen
Private tmpCount As Long

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
Dim m As String
  Me.Caption = "Info " & App.Title
  lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
  m = m & "Click Solve more times;" & vbCrLf
  m = m & "U: Undo; R: Redo; O: Open; S: SaveAs; I: Info;" & vbCrLf
  m = m & "The tooltips of Edits give info about next move;" & vbCrLf
  m = m & "Nr: Number of the cell lefttop=1, rightbottom=81" & vbCrLf
  m = m & "B: Number of quadratic block of cells lefttop=1, rightbottom=9" & vbCrLf
  m = m & "L: Line number firsttop=1, lastbottom=9" & vbCrLf
  m = m & "C: Column number left=1, right=9" & vbCrLf
  m = m & "PossV: possible value(s) of this cell" & vbCrLf
  lblDescription.Caption = lblDescription.Caption & vbCrLf & m
  tmpCount = 100
  PbMBOIngcom.MousePointer = 99
  PbMBOIngcom.MouseIcon = LoadResPicture(1, vbResCursor)
End Sub

Private Sub LblMBOINGCOM_Click()
  Call Shell("explorer.exe http://mbo-ing.com", vbMaximizedFocus)
End Sub

Private Sub PbMBOIngcom_Click()
  Call LblMBOINGCOM_Click
End Sub

Private Sub Timer1_Timer()
  tmpCount = tmpCount + 1
  If tmpCount > 105 Then tmpCount = 101
  Set ImgSunRun.Picture = LoadResPicture(tmpCount, vbResBitmap)
  'Image3.Refresh
End Sub
