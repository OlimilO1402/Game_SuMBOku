VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Image"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public MustInherit Class Image
'          Inherits System.MarshalByRefObject
'     Member von: System.Drawing
'Zusammenfassung:
' Eine abstrakte Basisklasse, die Funktionen für die von System.Drawing.Bitmap
' und System.Drawing.Imaging.Metafile abgeleiteten Klassen bereitstellt.
Option Explicit
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal X As Long, ByVal Y As Long, ByVal fStyle As Long) As Long
Private Declare Function ImageList_GetIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal Flags As Long) As Long
Private Declare Function DrawStateA Lib "user32" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal fuFlags As Long) As Long
Public Tag As String
Private mItem As Long
Private mhImageList As Long
Private mFlags As Long
'Private mHbmp As Long
'Private mHPal As Long
Private mPic As StdPicture
'

Private Sub Class_Initialize()
  'mFlags = &H3 Or &H20 'was sind das für Flags ?
  'mFlags = ILD_MASK Or ILD_TRANSPARENT Or ILD_OVERLAYMASK
End Sub
Private Sub Class_Terminate()
  Dispose
End Sub
Friend Sub Init(hImageList As Long, Item As Long, Optional aTag As String)
  mhImageList = hImageList
  mItem = Item
  Tag = aTag
End Sub

'Public Overridable NotOverridable Function Clone() As Object
'Public Overridable NotOverridable Sub Dispose()
Public Sub Dispose()
  'irgendwas löschen
End Sub
'von Graphics
Public Sub DrawImage(hhDC As OLE_HANDLE, Optional X As Long, Optional Y As Long, Optional Style As Long)
  Call ImageList_Draw(mhImageList, mItem - 1, hhDC, X, Y, mFlags) 'Style)
End Sub
Public Sub DrawDisabled(hhDC As Long, X As Long, Y As Long)
Dim hicon As Long
Dim hBrush As Long
  hicon = ImageList_GetIcon(mhImageList, mItem - 1, &H1)
  Call DrawStateA(hhDC, hBrush, 0, hicon, 0, X, Y, 0, 0, mFlags)
End Sub

'Public ReadOnly Property Flags() As Integer
Public Property Get Flags() As Long
  Flags = mFlags
End Property

'Public ReadOnly Property FrameDimensionsList() As System.Guid()

'Public Shared Function FromFile(ByVal filename As String) As System.Drawing.Image
'Public Shared Function FromFile(ByVal filename As String, ByVal useEmbeddedColorManagement As Boolean) As System.Drawing.Image
Public Function FromFile(ByVal filename As String, Optional ByVal useEmbeddedColorManagement As Boolean) As Image
  Set mPic = LoadPicture(filename)
End Function

'Public Shared Function FromHbitmap(ByVal hbitmap As System.IntPtr) As System.Drawing.Bitmap
'Public Shared Function FromHbitmap(ByVal hbitmap As System.IntPtr, ByVal hpalette As System.IntPtr) As System.Drawing.Bitmap
'Public Function FromHbitmap(ByVal hbitmap As Long, Optional ByVal hpalette As Long) As Image
''  mPic.Handle = hbitmap
''  mPic.hPal = hpalette
'  Set FromHbitmap = Me
'End Function

Public Function FromStdPicture(aPic As StdPicture) As Image
  Set mPic = aPic
  Set FromStdPicture = Me
End Function

Public Function Picture() As StdPicture
  Set Picture = mPic
End Function
'Public Shared Function FromStream(ByVal stream As System.IO.Stream) As System.Drawing.Image
'Public Shared Function FromStream(ByVal stream As System.IO.Stream, ByVal useEmbeddedColorManagement As Boolean) As System.Drawing.Image
'Public Shared Function FromStream(ByVal stream As System.IO.Stream, ByVal useEmbeddedColorManagement As Boolean, ByVal validateImageData As Boolean) As System.Drawing.Image

'Public Function GetBounds(ByRef pageUnit As System.Drawing.GraphicsUnit) As System.Drawing.RectangleF

'Public Function GetEncoderParameterList(ByVal encoder As System.Guid) As System.Drawing.Imaging.EncoderParameters
'Public Function GetFrameCount(ByVal dimension As System.Drawing.Imaging.FrameDimension) As Integer
'Public Shared Function GetPixelFormatSize(ByVal pixfmt As System.Drawing.Imaging.PixelFormat) As Integer
'Public Function GetPropertyItem(ByVal propid As Integer) As System.Drawing.Imaging.PropertyItem
'Public Function GetThumbnailImage(ByVal thumbWidth As Integer, ByVal thumbHeight As Integer, ByVal callback As System.Drawing.Image.GetThumbnailImageAbort, ByVal callbackData As System.IntPtr) As System.Drawing.Image
'Public ReadOnly Property Height() As Integer
'Public ReadOnly Property HorizontalResolution() As Single
'Public Shared Function IsAlphaPixelFormat(ByVal pixfmt As System.Drawing.Imaging.PixelFormat) As Boolean
'Public Shared Function IsCanonicalPixelFormat(ByVal pixfmt As System.Drawing.Imaging.PixelFormat) As Boolean
'Public Shared Function IsExtendedPixelFormat(ByVal pixfmt As System.Drawing.Imaging.PixelFormat) As Boolean
'Public Property Palette() As System.Drawing.Imaging.ColorPalette
'Public ReadOnly Property PhysicalDimension() As System.Drawing.SizeF
'Public ReadOnly Property PixelFormat() As System.Drawing.Imaging.PixelFormat
'Public ReadOnly Property PropertyIdList() As Integer()
'Public ReadOnly Property PropertyItems() As System.Drawing.Imaging.PropertyItem()
'Public ReadOnly Property RawFormat() As System.Drawing.Imaging.ImageFormat
'Public Sub RemovePropertyItem(ByVal propid As Integer)
'Public Sub RotateFlip(ByVal rotateFlipType As System.Drawing.RotateFlipType)
'Public Sub Save(ByVal filename As String)
'Public Sub Save(ByVal filename As String, ByVal encoder As System.Drawing.Imaging.ImageCodecInfo, ByVal encoderParams As System.Drawing.Imaging.EncoderParameters)
'Public Sub Save(ByVal filename As String, ByVal format As System.Drawing.Imaging.ImageFormat)
'Public Sub Save(ByVal stream As System.IO.Stream, ByVal encoder As System.Drawing.Imaging.ImageCodecInfo, ByVal encoderParams As System.Drawing.Imaging.EncoderParameters)
'Public Sub Save(ByVal stream As System.IO.Stream, ByVal format As System.Drawing.Imaging.ImageFormat)
'Public Sub SaveAdd(ByVal image As System.Drawing.Image, ByVal encoderParams As System.Drawing.Imaging.EncoderParameters)
'Public Sub SaveAdd(ByVal encoderParams As System.Drawing.Imaging.EncoderParameters)
'Public Function SelectActiveFrame(ByVal dimension As System.Drawing.Imaging.FrameDimension, ByVal frameIndex As Integer) As Integer
'Public Sub SetPropertyItem(ByVal propitem As System.Drawing.Imaging.PropertyItem)
'Public ReadOnly Property Size() As System.Drawing.Size
'Public ReadOnly Property VerticalResolution() As Single
'Public ReadOnly Property Width() As Integer
'Protected Overridable Sub Dispose(ByVal disposing As Boolean)
'Protected Overrides Sub Finalize()

'Public Function ExtractIcon() As Long
''Dim hIcon As Long
'  ExtractIcon = ImageList_GetIcon(mhImageList, mItem - 1, &H1)
'End Function

'Dim img As System.Drawing.Image
'System.Drawing.Image.Clone()
'System.Drawing.Image.CreateObjRef()
'System.Drawing.Image.Dispose()
'System.Drawing.Image.Flags()
'System.Drawing.Image.FrameDimensionsList()
'System.Drawing.Image.FromFile()
'System.Drawing.Image.FromHbitmap()
'System.Drawing.Image.FromStream()
'System.Drawing.Image.GetBounds() As RectangleF
'System.Drawing.Image.GetEncoderParameterList()
'System.Drawing.Image.GetFrameCount()
'System.Drawing.Image.GetLifetimeService()
'System.Drawing.Image.GetPixelFormatSize()
'System.Drawing.Image.GetPropertyItem()
'System.Drawing.Image.GetThumbnailImage()
'System.Drawing.Image.GetThumbnailImageAbort()
'System.Drawing.Image.GetType()
'System.Drawing.Image.Height()
'System.Drawing.Image.HorizontalResolution()
'System.Drawing.Image.InitializeLifetimeService()
'System.Drawing.Image.IsAlphaPixelFormat()
'System.Drawing.Image.IsCanonicalPixelFormat()
'System.Drawing.Image.IsExtendedPixelFormat()
'System.Drawing.Image.Palette() as System.Drawing.Imaging.ColorPalette
'System.Drawing.Image.PhysicalDimension() 'as SizeF
'System.Drawing.Image.PixelFormat()
'System.Drawing.Image.PropertyIdList()
'System.Drawing.Image.PropertyItems()
'System.Drawing.Image.RawFormat()
'System.Drawing.Image.RemovePropertyItem()
'System.Drawing.Image.RotateFlip()
'System.Drawing.Image.Save()
'System.Drawing.Image.SaveAdd()
'System.Drawing.Image.SelectActiveFrame()
'System.Drawing.Image.SetPropertyItem()
'System.Drawing.Image.Size()
'System.Drawing.Image.VerticalResolution()
'System.Drawing.Image.Width()

