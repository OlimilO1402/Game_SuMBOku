VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DynTextBoxes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mParentForm As Form
Private mCol As Collection
Private Const TBoxName As String = "TBoxName"
Public Event Change(index As Long)
Public Event Click(index As Long)
Public Event DblClick(index As Long)
Public Event DragDrop(index As Long, Source As Control, X As Single, Y As Single)
Public Event DragOver(index As Long, Source As Control, X As Single, Y As Single, State As Integer)
Public Event GotFocus(index As Long)
Public Event KeyDown(index As Long, KeyCode As Integer, Shift As Integer)
Public Event KeyPress(index As Long, KeyAscii As Integer)
Public Event KeyUp(index As Long, KeyCode As Integer, Shift As Integer)
Public Event LinkClose(index As Long)
Public Event LinkError(index As Long, LinkErr As Integer)
Public Event LinkNotify(index As Long)
Public Event LinkOpen(index As Long, Cancel As Integer)
Public Event LostFocus(index As Long)
Public Event MouseDown(index As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(index As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(index As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event OLECompleteDrag(index As Long, Effect As Long)
Public Event OLEDragDrop(index As Long, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event OLEDragOver(index As Long, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Public Event OLEGiveFeedback(index As Long, Effect As Long, DefaultCursors As Boolean)
Public Event OLESetData(index As Long, Data As DataObject, DataFormat As Integer)
Public Event OLEStartDrag(index As Long, Data As DataObject, AllowedEffects As Long)
Public Event Validate(CIndex As Long, ancel As Boolean)

Private Sub Class_Initialize()
  Set mCol = New Collection
End Sub
Private Sub Class_Terminate()
  Set mCol = Nothing
End Sub

Public Sub NewC(ParentForm As Form)
  Set mParentForm = ParentForm
End Sub

Public Function Add() As DynTextBox
  Dim i As Long: i = mCol.Count + 1
  Set Add = New_DynTextBox(mParentForm, Me, TBoxName & CStr(i), i)
  Call mCol.Add(Add)
End Function

Public Property Get Count() As Long
  Count = mCol.Count
End Property

Public Property Set Font(aFnt As StdFont)
Dim aTBox As DynTextBox
  For Each aTBox In mCol
    Set aTBox.TextBox.Font = aFnt
  Next
End Property
Public Property Get TextBox(index As Long) As TextBox
  Dim aTBox As DynTextBox: Set aTBox = mCol.Item(index)
  Set TextBox = aTBox.TextBox
End Property

'##############################'  RaiseEvents  '##############################'
Friend Sub OnChange(index As Long):   RaiseEvent Change(index): End Sub
Friend Sub OnClick(index As Long):    RaiseEvent Click(index):  End Sub
Friend Sub OnDblClick(index As Long): RaiseEvent DblClick(index): End Sub
Friend Sub OnDragDrop(index As Long, Source As Control, X As Single, Y As Single):                   RaiseEvent DragDrop(index, Source, X, Y): End Sub
Friend Sub OnDragOver(index As Long, Source As Control, X As Single, Y As Single, State As Integer): RaiseEvent DragOver(index, Source, X, Y, State): End Sub
Friend Sub OnGotFocus(index As Long): RaiseEvent GotFocus(index): End Sub
Friend Sub OnKeyDown(index As Long, KeyCode As Integer, Shift As Integer): RaiseEvent KeyDown(index, KeyCode, Shift): End Sub
Friend Sub OnKeyPress(index As Long, KeyAscii As Integer):                 RaiseEvent KeyPress(index, KeyAscii): End Sub
Friend Sub OnKeyUp(index As Long, KeyCode As Integer, Shift As Integer):   RaiseEvent KeyUp(index, KeyCode, Shift): End Sub
Friend Sub OnLinkClose(index As Long): RaiseEvent LinkClose(index): End Sub
Friend Sub OnLinkError(index As Long, LinkErr As Integer): RaiseEvent LinkError(index, LinkErr): End Sub
Friend Sub OnLinkNotify(index As Long):                    RaiseEvent LinkNotify(index):         End Sub
Friend Sub OnLinkOpen(index As Long, Cancel As Integer):   RaiseEvent LinkOpen(index, Cancel):   End Sub
Friend Sub OnLostFocus(index As Long):                     RaiseEvent LostFocus(index):          End Sub
Friend Sub OnMouseDown(index As Long, Button As Integer, Shift As Integer, X As Single, Y As Single): RaiseEvent MouseDown(index, Button, Shift, X, Y): End Sub
Friend Sub OnMouseMove(index As Long, Button As Integer, Shift As Integer, X As Single, Y As Single): RaiseEvent MouseMove(index, Button, Shift, X, Y): End Sub
Friend Sub OnMouseUp(index As Long, Button As Integer, Shift As Integer, X As Single, Y As Single):   RaiseEvent MouseUp(index, Button, Shift, X, Y):   End Sub
Friend Sub OnOLECompleteDrag(index As Long, Effect As Long): RaiseEvent OLECompleteDrag(index, Effect): End Sub
Friend Sub OnOLEDragDrop(index As Long, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single):                   RaiseEvent OLEDragDrop(index, Data, Effect, Button, Shift, X, Y):         End Sub
Friend Sub OnOLEDragOver(index As Long, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer): RaiseEvent OLEDragOver(index, Data, Effect, Button, Shift, X, Y, State):  End Sub
Friend Sub OnOLEGiveFeedback(index As Long, Effect As Long, DefaultCursors As Boolean):                                                                       RaiseEvent OLEGiveFeedback(index, Effect, DefaultCursors):                End Sub
Friend Sub OnOLESetData(index As Long, Data As DataObject, DataFormat As Integer):                                                                            RaiseEvent OLESetData(index, Data, DataFormat):                           End Sub
Friend Sub OnOLEStartDrag(index As Long, Data As DataObject, AllowedEffects As Long):                                                                         RaiseEvent OLEStartDrag(index, Data, AllowedEffects):                     End Sub
Friend Sub OnValidate(index As Long, Cancel As Boolean): RaiseEvent Validate(index, Cancel): End Sub

