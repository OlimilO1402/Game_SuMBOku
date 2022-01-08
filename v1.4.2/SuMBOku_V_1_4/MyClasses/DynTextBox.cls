VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DynTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mTB As TextBox
Attribute mTB.VB_VarHelpID = -1
Private mIndex As Long
Private mTBCol As DynTextBoxes

Private Sub Class_Initialize()
  '
End Sub
Private Sub Class_Terminate()
  Set mTB = Nothing
End Sub

Public Sub NewC(aFrm As Form, aTBCol As DynTextBoxes, aStrName As String, Index As Long)
  mIndex = Index
  Set mTBCol = aTBCol
  Set mTB = aFrm.Controls.Add("VB.TextBox", aStrName & CStr(mIndex), aFrm)
  mTB.Visible = True
  Call mTB.Move(120, 120, 1215, 285)
End Sub

Public Property Get TextBox() As TextBox
  Set TextBox = mTB
End Property

'##############################'  mTB Events '##############################'
Private Sub mTB_Change():   Call mTBCol.OnChange(mIndex):   End Sub
Private Sub mTB_Click():    Call mTBCol.OnClick(mIndex):    End Sub
Private Sub mTB_DblClick(): Call mTBCol.OnDblClick(mIndex): End Sub
Private Sub mTB_DragDrop(Source As Control, X As Single, Y As Single):                   Call mTBCol.OnDragDrop(mIndex, Source, X, Y):        End Sub
Private Sub mTB_DragOver(Source As Control, X As Single, Y As Single, State As Integer): Call mTBCol.OnDragOver(mIndex, Source, X, Y, State): End Sub
Private Sub mTB_GotFocus(): Call mTBCol.OnGotFocus(mIndex):  End Sub
Private Sub mTB_KeyDown(KeyCode As Integer, Shift As Integer): Call mTBCol.OnKeyDown(mIndex, KeyCode, Shift): End Sub
Private Sub mTB_KeyPress(KeyAscii As Integer):                 Call mTBCol.OnKeyPress(mIndex, KeyAscii):      End Sub
Private Sub mTB_KeyUp(KeyCode As Integer, Shift As Integer):   Call mTBCol.OnKeyUp(mIndex, KeyCode, Shift):   End Sub
Private Sub mTB_LinkClose():                   Call mTBCol.OnLinkClose(mIndex):          End Sub
Private Sub mTB_LinkError(LinkErr As Integer): Call mTBCol.OnLinkError(mIndex, LinkErr): End Sub
Private Sub mTB_LinkNotify():                  Call mTBCol.OnLinkNotify(mIndex):         End Sub
Private Sub mTB_LinkOpen(Cancel As Integer):   Call mTBCol.OnLinkOpen(mIndex, Cancel):   End Sub
Private Sub mTB_LostFocus(): Call mTBCol.OnLostFocus(mIndex): End Sub
Private Sub mTB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single): Call mTBCol.OnMouseDown(mIndex, Button, Shift, X, Y): End Sub
Private Sub mTB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single): Call mTBCol.OnMouseMove(mIndex, Button, Shift, X, Y): End Sub
Private Sub mTB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single):   Call mTBCol.OnMouseUp(mIndex, Button, Shift, X, Y):   End Sub
Private Sub mTB_OLECompleteDrag(Effect As Long):                                                                                                  Call mTBCol.OnOLECompleteDrag(mIndex, Effect):                               End Sub
Private Sub mTB_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single):                   Call mTBCol.OnOLEDragDrop(mIndex, Data, Effect, Button, Shift, X, Y):        End Sub
Private Sub mTB_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer): Call mTBCol.OnOLEDragOver(mIndex, Data, Effect, Button, Shift, X, Y, State): End Sub
Private Sub mTB_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean):                                                                       Call mTBCol.OnOLEGiveFeedback(mIndex, Effect, DefaultCursors):               End Sub
Private Sub mTB_OLESetData(Data As DataObject, DataFormat As Integer):                                                                            Call mTBCol.OnOLESetData(mIndex, Data, DataFormat):                          End Sub
Private Sub mTB_OLEStartDrag(Data As DataObject, AllowedEffects As Long):                                                                         Call mTBCol.OnOLEStartDrag(mIndex, Data, AllowedEffects):                    End Sub
Private Sub mTB_Validate(Cancel As Boolean): Call mTBCol.OnValidate(mIndex, Cancel): End Sub
