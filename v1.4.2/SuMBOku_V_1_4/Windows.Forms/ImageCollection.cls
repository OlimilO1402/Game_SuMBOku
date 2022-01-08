VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ImageCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public NotInheritable Class ImageList.ImageCollection
'          Inherits System.Object
'     Member von: System.Windows.Forms
'Zusammenfassung:
' Kapselt die Auflistung von System.Drawing.Image-Objekten in eine
' System.Windows.Forms.ImageList.
Option Explicit
Private Declare Function ImageList_AddIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal hicon As Long) As Long
Private Declare Function ImageList_AddMasked Lib "comctl32.dll" (ByVal himl As Long, ByVal hbmImage As Long, ByVal crMask As Long) As Long
Private Declare Function ImageList_Remove Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long) As Long
Private mCol As New Collection
Private mhImageList As Long
Private mTransparentColor As Long

Private Sub Class_Initialize()
  Call Clear
End Sub
Private Sub Class_Terminate()
  Set mCol = Nothing
End Sub

Friend Sub Init(hImageList As Long, aTransparentColor As Long)
  mhImageList = hImageList
  mTransparentColor = aTransparentColor
End Sub

'Public Sub Add(ByVal value As System.Drawing.Icon)
'Public Sub Add(ByVal value As System.Drawing.Image)
'Public Function Add(ByVal value As System.Drawing.Image, ByVal transparentColor As System.Drawing.Color) As Integer
Public Function Add(value As image, Optional ByVal aTransparentColor As Long = -1, Optional aTag As String) As Long
  mCol.Add value
  Call value.Init(mhImageList, mCol.Count, aTag)
  If aTransparentColor = -1 Then aTransparentColor = mTransparentColor
  Call ImageList_AddMasked(mhImageList, value.Picture.Handle, aTransparentColor)
End Function
'Public Sub AddIcon(Icon As Long)
'Dim NewIcon As New image
'  mCol.Add NewIcon
'  Call NewIcon.Init(mhImageList, mCol.Count)
'  Call ImageList_AddIcon(mhImageList, Icon)
'End Sub
'Public Function AddStrip(ByVal value As System.Drawing.Image) As Integer

'Public Overridable NotOverridable Sub Clear()
Public Sub Clear()
  Set mCol = New Collection
End Sub

'Public Function Contains(ByVal image As System.Drawing.Image) As Boolean

'Public Overridable NotOverridable ReadOnly Property Count() As Integer
Public Property Get Count() As Long
  Count = mCol.Count
End Property

'Public ReadOnly Property Empty() As Boolean
Public Property Get EEmpty() As Boolean
  EEmpty = (mCol.Count = 0)
End Property

'Public Overridable NotOverridable Function GetEnumerator() As System.Collections.IEnumerator
Public Function GetEnumerator() As IUnknown
Attribute GetEnumerator.VB_UserMemId = -4
  Set GetEnumerator = mCol.[_NewEnum]
End Function

'Public Function IndexOf(ByVal image As System.Drawing.Image) As Integer
Public Function IndexOf(ByVal img As image) As Long
Dim i As Long, im As image
  For i = 1 To mCol.Count
    If im Is img Then
      IndexOf = i
      Exit Function
    End If
  Next
End Function

'Public Default Property Item(ByVal index As Integer) As System.Drawing.Image
Public Property Get Item(Optional index As Long = 1) As image 'System.Drawing.Image
  Set Item = mCol(index)
  Call Item.Init(mhImageList, index)
End Property


'heiﬂt unter dotNET ImageCollection
'im ilC As New ImageList.ImageCollection
'ImageList.ImageCollection.Add()
'ImageList.ImageCollection.AddStrip()
'ImageList.ImageCollection.Clear()
'ImageList.ImageCollection.Contains()
'ImageList.ImageCollection.Count()
'ImageList.ImageCollection.Empty()
'ImageList.ImageCollection.GetEnumerator()
'ImageList.ImageCollection.GetType()
'ImageList.ImageCollection.IndexOf()
'ImageList.ImageCollection.IsReadOnly()
'ImageList.ImageCollection.Item()
'ImageList.ImageCollection.Remove()
'ImageList.ImageCollection.RemoveAt()

