VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ImageList"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public NotInheritable Class ImageList
'          Inherits System.ComponentModel.Component
'     Member von: System.Windows.Forms
'Zusammenfassung:
' Stellt Methoden zur Verwaltung einer Auflistung von System.Drawing.Image-Objekten
' bereit. Die Klasse kann nicht vererbt werden.
Option Explicit
Private Type IMAGELISTDRAWPARAMS
    cbSize  As Long
    himl    As Long
    i       As Long
    hhdcDst As Long 'hDC
    X       As Long
    Y       As Long
    cx      As Long
    cy      As Long
    xBitmap As Long '// x offest from the upperleft of bitmap
    yBitmap As Long '// y offset from the upperleft of bitmap
    rgbBk   As Long   'COLORREF;
    rgbFg   As Long   'COLORREF;
    fStyle  As Long  'UINT;
    dwRop   As Long   'DWORD;
End Type
'FlagConstanten:
Private Const CLR_NONE = &HFFFFFFFF
Private Const CLR_DEFAULT = &HFF000000

Private Const ILC_COLOR As Long = &H0
Private Const ILC_MASK As Long = &H1
Private Const ILC_COLOR4 As Long = &H4
Private Const ILC_COLOR8 As Long = &H8
Private Const ILC_COLOR16 As Long = &H10
Private Const ILC_COLOR24 As Long = &H18
Private Const ILC_COLOR32 As Long = &H20
Private Const ILC_COLORDDB As Long = &HFE&
Private Const ILC_PALETTE As Long = &H800

Private Declare Function ImageList_Add Lib "comctl32.dll" (ByVal himl As Long, ByVal hbmImage As Long, ByVal hbmMask As Long) As Long
Private Declare Function ImageList_AddMasked Lib "comctl32.dll" (ByVal himl As Long, ByVal hbmImage As Long, ByVal crMask As Long) As Long
Private Declare Function ImageList_BeginDrag Lib "comctl32.dll" (ByVal himlTrack As Long, ByVal iTrack As Long, ByVal dxHotspot As Long, ByVal dyHotspot As Long) As Long
Private Declare Function ImageList_Copy Lib "comctl32.dll" (ByVal himlDst As Long, ByVal iDst As Long, ByVal himlSrc As Long, ByVal iSrc As Long, ByVal uFlags As Long) As Long
Private Declare Function ImageList_Create Lib "comctl32.dll" (ByVal cx As Long, ByVal cy As Long, ByVal Flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
Private Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal himl As Long) As Long
Private Declare Function ImageList_DragEnter Lib "comctl32.dll" (ByVal hwndLock As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ImageList_DragLeave Lib "comctl32.dll" (ByVal hwndLock As Long) As Long
Private Declare Function ImageList_DragMove Lib "comctl32.dll" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ImageList_DragShowNolock Lib "comctl32.dll" (ByVal fShow As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal X As Long, ByVal Y As Long, ByVal fStyle As Long) As Long
Private Declare Function ImageList_DrawEx Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal rgbBk As Long, ByVal rgbFg As Long, ByVal fStyle As Long) As Long
Private Declare Function ImageList_DrawIndirect Lib "comctl32.dll" (ByRef pimldp As IMAGELISTDRAWPARAMS) As Long
Private Declare Function ImageList_Duplicate Lib "comctl32.dll" (ByVal himl As Long) As Long
Private Declare Sub ImageList_EndDrag Lib "comctl32.dll" ()
Private Declare Function ImageList_GetBkColor Lib "comctl32.dll" (ByVal himl As Long) As Long
'Private Declare Function ImageList_GetDragImage Lib "comctl32.dll" (ByRef ppt As Point, ByRef pptHotspot As Point) As Long
Private Declare Function ImageList_GetIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal Flags As Long) As Long
Private Declare Function ImageList_GetIconSize Lib "comctl32.dll" (ByVal himl As Long, ByRef cx As Long, ByRef cy As Long) As Long
Private Declare Function ImageList_GetImageCount Lib "comctl32.dll" (ByVal himl As Long) As Long
'Private Declare Function ImageList_GetImageInfo Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByRef pImageInfo As IMAGEINFO) As Long
Private Declare Function ImageList_LoadImage Lib "comctl32.dll" (ByVal hi As Long, ByVal lpbmp As String, ByVal cx As Long, ByVal cGrow As Long, ByVal crMask As Long, ByVal uType As Long, ByVal uFlags As Long) As Long
Private Declare Function ImageList_Merge Lib "comctl32.dll" (ByVal himl1 As Long, ByVal i1 As Long, ByVal himl2 As Long, ByVal i2 As Long, ByVal dx As Long, ByVal dy As Long) As Long
Private Declare Function ImageList_Read Lib "comctl32.dll" (ByRef pstm As Long) As Long
Private Declare Function ImageList_Remove Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long) As Long
Private Declare Function ImageList_Replace Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hbmImage As Long, ByVal hbmMask As Long) As Long
Private Declare Function ImageList_ReplaceIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hicon As Long) As Long
Private Declare Function ImageList_SetBkColor Lib "comctl32.dll" (ByVal himl As Long, ByVal clrBk As Long) As Long
Private Declare Function ImageList_SetDragCursorImage Lib "comctl32.dll" (ByVal himlDrag As Long, ByVal iDrag As Long, ByVal dxHotspot As Long, ByVal dyHotspot As Long) As Long
Private Declare Function ImageList_SetIconSize Lib "comctl32.dll" (ByVal himl As Long, ByVal cx As Long, ByVal cy As Long) As Long
Private Declare Function ImageList_SetImageCount Lib "comctl32.dll" (ByVal himl As Long, ByVal uNewCount As Long) As Long
Private Declare Function ImageList_SetOverlayImage Lib "comctl32.dll" (ByVal himl As Long, ByVal iImage As Long, ByVal iOverlay As Long) As Long
Private Declare Function ImageList_Write Lib "comctl32.dll" (ByVal himl As Long, ByRef pstm As Long) As Long
Private mHandle As Long
Private mTransparentColor As Long
Private mImages As New ImageCollection
Private mImageW As Long  'eigentlich ein StructType Size
Private mImageH As Long
Private mFlags As Long
'---------------
Private Const ILD_NORMAL As Long = &H0
Private Const ILD_TRANSPARENT As Long = &H1
Private Const ILD_BLEND25 As Long = &H2
Private Const ILD_BLEND50 As Long = &H4
Private Const ILD_FOCUS As Long = ILD_BLEND25
Private Const ILD_BLEND As Long = ILD_BLEND50
Private Const ILD_MASK As Long = &H10
Private Const ILD_IMAGE As Long = &H20
Private Const ILD_ROP As Long = &H40
Private Const ILD_OVERLAYMASK As Long = &HF00&
Private Const ILD_SELECTED As Long = ILD_BLEND50
Public Enum ColorDepth
  Depth4Bit = 4   '=ILC_COLOR4
  Depth8Bit = 8   '=ILC_COLOR8
  Depth16Bit = 16 '=ILC_COLOR16
  Depth24Bit = 24 '=ILC_COLOR24
  Depth32Bit = 32 '=ILC_COLOR32
End Enum
Private mColorDepth As ColorDepth

Private Sub Class_Initialize()
'                            cx, cy, Flags, cInitial, cGrow
  mColorDepth = Depth16Bit
  mFlags = ILC_MASK Or mColorDepth 'ILC_COLORDDB
  mImageW = 16: mImageH = 16
End Sub

Private Sub Class_Terminate()
  Call Dispose(True)
End Sub

'Protected Overrides Sub Dispose(ByVal disposing As Boolean)
Public Sub Dispose(ByVal disposing As Boolean)
  Call ImageList_Destroy(mHandle)
End Sub

'Public Property ColorDepth() As System.Windows.Forms.ColorDepth
Public Property Get ColorDepth() As ColorDepth
  ColorDepth = mColorDepth
End Property
Public Property Let ColorDepth(LngVal As ColorDepth)
  mColorDepth = LngVal
End Property


'Public Sub Draw(ByVal g As System.Drawing.Graphics, ByVal x As Integer, ByVal y As Integer, ByVal index As Integer)
'Public Sub Draw(ByVal g As System.Drawing.Graphics, ByVal x As Integer, ByVal y As Integer, ByVal width As Integer, ByVal height As Integer, ByVal index As Integer)
Public Sub Draw(ByVal hhDC As Long, ByVal X As Long, ByVal Y As Long, Optional ByVal Width As Long, Optional ByVal Height As Long, Optional ByVal index As Long = 1)
'ImageList_Draw         (ByVal himl As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal x As Long, ByVal y As Long, ByVal fStyle As Long) As Long
'ImageList_DrawEx       (ByVal himl As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal rgbBk As Long, ByVal rgbFg As Long, ByVal fStyle As Long) As Long
'ImageList_DrawIndirect (ByRef pimldp As IMAGELISTDRAWPARAMS) As Long
Dim rv As Long
Dim mFlg As Long
Dim hhnd As Long
  mFlg = ILD_MASK Or ILD_TRANSPARENT Or ILD_OVERLAYMASK
  hhnd = Handle
  rv = ImageList_Draw(hhnd, index, hhDC, X, Y, mFlg)
  If rv <> 1 Then MsgBox "Fehler in ImageList_Draw."
End Sub

'Public ReadOnly Property Handle() As System.IntPtr
Public Property Get Handle() As Long
  If HandleCreated Then
    '
  Else
    mHandle = ImageList_Create(mImageW, mImageH, mFlags, 0&, 0&)
  End If
  Handle = mHandle
End Property
'Public ReadOnly Property HandleCreated() As Boolean
Public Property Get HandleCreated() As Boolean
  HandleCreated = (mHandle <> 0)
End Property

'Public ReadOnly Property Images() As System.Windows.Forms.ImageList.ImageCollection
Public Property Get Images() As ImageCollection
  Call mImages.Init(Handle, mTransparentColor)
  Set Images = mImages
End Property

'Public Property ImageSize() As System.Drawing.Size
Public Property Get ImageSizeWidth() As Long
  ImageSizeWidth = mImageW
End Property
Public Property Let ImageSizeWidth(LngVal As Long)
  mImageW = LngVal
End Property
Public Property Get ImageSizeHeight() As Long
  ImageSizeHeight = mImageH
End Property
Public Property Let ImageSizeHeight(LngVal As Long)
  mImageH = LngVal
End Property

'Public Property ImageStream() As System.Windows.Forms.ImageListStreamer
'Public Sub New()
'Public Sub New(ByVal container As System.ComponentModel.IContainer)


'Public Overrides Function ToString() As String
Public Function ToString() As String
  ToString = "[System.Windows.Forms.ImageList] Images.Count: " & CStr(Images.Count) & ", ImageSize: {Width=16, Height =16} "
End Function

'Public Property TransparentColor() As System.Drawing.Color
Public Property Get TransparentColor() As OLE_COLOR
  TransparentColor = mTransparentColor
End Property
Public Property Let TransparentColor(ClrVal As OLE_COLOR)
  mTransparentColor = ClrVal
End Property


'in der dotNET-Klasse ImageList gibt es folgende Public Member:
    'ImageList1.ColorDepth()
    'ImageList1.Container
    'ImageList1.CreateObjRef()
    'ImageList1.Dispose()
    'ImageList1.GetLifetimeService()
    'ImageList1.GetType()
    'ImageList1.Handle()
    'ImageList1.HandleCreated()
    'ImageList1.ImageCollection() 'die Klasse
    'ImageList1.Images()          'ein Objekt von ImageCollection
    'ImageList1.ImageSize()
    'ImageList1.ImageStream()
    'ImageList1.InitializeLifetimeService()
    'ImageList1.Site()
    'ImageList1.ToString()
    'ImageList1.TransparentColor()
'im Objektbrowser
'Public Property ColorDepth() As System.Windows.Forms.ColorDepth
'Public Sub Draw(ByVal g As System.Drawing.Graphics, ByVal x As Integer, ByVal y As Integer, ByVal index As Integer)
'Public Sub Draw(ByVal g As System.Drawing.Graphics, ByVal x As Integer, ByVal y As Integer, ByVal width As Integer, ByVal height As Integer, ByVal index As Integer)
'Public Sub Draw(ByVal g As System.Drawing.Graphics, ByVal pt As System.Drawing.Point, ByVal index As Integer)
'Public ReadOnly Property Handle() As System.IntPtr
'Public ReadOnly Property HandleCreated() As Boolean
'Public ReadOnly Property Images() As System.Windows.Forms.ImageList.ImageCollection
'Public Property ImageSize() As System.Drawing.Size
'Public Property ImageStream() As System.Windows.Forms.ImageListStreamer
'Public Sub New()
'Public Sub New(ByVal container As System.ComponentModel.IContainer)
'Public Overrides Function ToString() As String
'Public Property TransparentColor() As System.Drawing.Color
'Protected Overrides Sub Dispose(ByVal disposing As Boolean)

'ListImages heiﬂt ImageCollection und
'ListImage  heiﬂt System.Drawing.Image

