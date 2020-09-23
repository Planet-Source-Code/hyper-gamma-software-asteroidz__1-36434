VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_Bitmap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    
  Private Const SRCAND = &H8800C6
  Private Const SRCCOPY = &HCC0020
  Private Const SRCERASE = &H440328
  Private Const SRCINVERT = &H660046
  Private Const SRCPAINT = &HEE0086
  
  Private Const IMAGE_BITMAP = &O0
  Private Const LR_LOADFROMFILE = 16
  
  Private Type Bitmap
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
  End Type
  
  Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
  Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
  Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
  Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
  Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
  Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
  Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
  Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
  Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
  Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
  Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
  Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long

  Private obj_Parent          As cls_Bitmaps
  
  Private lng_ImageDC         As Long
  Private lng_InvertImageDC   As Long
  Private lng_MaskDC          As Long
  
  Private lng_ImageOld        As Long
  Private lng_InvertImageOld  As Long
  Private lng_MaskOld         As Long
  
  Private lng_Width           As Long
  Private lng_Height          As Long
  Private lng_InvertImage     As Long
  
  Private lng_Mask            As Long
  
  Private bln_BitmapLoaded    As Boolean
  
  Private int_FramesWide      As Integer
  Private int_FramesHigh      As Integer
  
  Private int_FrameCount      As Integer
  Private int_FrameWidth      As Integer
  Private int_FrameHeight     As Integer
  '

Property Set Parent(obj_NewValue As cls_Bitmaps)
  
  Set obj_Parent = obj_NewValue
  
End Property

Property Get Parent() As cls_Bitmaps
  
  Set Parent = obj_Parent
  
End Property

Property Get ImageDC() As Long
     
  ImageDC = lng_ImageDC
                                   
End Property

Property Get InvertImageDC() As Long
  
  InvertImageDC = lng_InvertImageDC
                                   
End Property

Property Get MaskDC() As Long
  
  MaskDC = lng_MaskDC
                                   
End Property
  
Property Let FramesWide(int_NewValue As Integer)

  int_FramesWide = int_NewValue
  int_FrameWidth = lng_Width \ int_FramesWide
        
End Property

Property Get FramesWide() As Integer
  
  FramesWide = int_FramesWide
  
End Property

Property Let FramesHigh(int_NewValue As Integer)

  int_FramesHigh = int_NewValue
  int_FrameHeight = lng_Height \ int_FramesHigh
    
End Property

Property Get FramesHigh() As Integer
  
  FramesHigh = int_FramesHigh
  
End Property

Property Get FrameCount() As Integer
  
  FrameCount = int_FrameCount
  
End Property

Property Get FrameWidth() As Integer
  
  FrameWidth = int_FrameWidth
  
End Property

Property Get FrameHeight() As Integer
  
  FrameHeight = int_FrameHeight
  
End Property

Property Get Width() As Long

  Width = lng_Width
  
End Property

Property Get Height() As Long

  Height = lng_Height
  
End Property

Private Sub CleanUp()

  Dim lng_Result As Long
  
  lng_Result = DeleteObject(SelectObject(lng_ImageDC, lng_ImageOld))
  lng_Result = DeleteObject(SelectObject(lng_InvertImageDC, lng_InvertImageOld))
  lng_Result = DeleteObject(SelectObject(lng_MaskDC, lng_MaskOld))
    
  lng_Result = DeleteDC(lng_ImageDC)
  lng_Result = DeleteDC(lng_InvertImageDC)
  lng_Result = DeleteDC(lng_MaskDC)
  
End Sub

Public Function Cls() As Boolean
  
  Dim lng_Result As Long
  
  lng_Result = BitBlt(lng_ImageDC, 0, 0, lng_Width, lng_Height, lng_ImageDC, 0, 0, SRCERASE)
  If (lng_Result = 0) Then Cls = False

End Function

Public Function BlitFast(lng_DestDC As Long, int_Left As Integer, int_Top As Integer, int_Frame As Integer, Optional bln_BlitTransParent As Boolean = True) As Boolean
    
  Dim int_SrcLeft As Integer
  Dim int_SrcTop  As Integer
  Dim lnt_Result  As Long
  
  On Error GoTo lbl_Abort
  
  BlitFast = True
  
  int_SrcLeft = int_FrameWidth * (int_Frame Mod int_FramesWide)
  int_SrcTop = IIf(int_FrameWidth = 1, 0, int_FrameHeight * (int_Frame \ FramesWide))
        
  If bln_BlitTransParent Then
  
    lnt_Result = BitBlt(lng_DestDC, int_Left, int_Top, int_FrameWidth, int_FrameHeight, lng_MaskDC, int_SrcLeft, int_SrcTop, SRCAND)
    If (lnt_Result = 0) Then GoTo lbl_Abort
    
    lnt_Result = BitBlt(lng_DestDC, int_Left, int_Top, int_FrameWidth, int_FrameHeight, lng_InvertImageDC, int_SrcLeft, int_SrcTop, SRCPAINT)
    If (lnt_Result = 0) Then GoTo lbl_Abort
  
  Else
    
    lnt_Result = BitBlt(lng_DestDC, int_Left, int_Top, int_FrameWidth, int_FrameHeight, lng_ImageDC, int_SrcLeft, int_SrcTop, SRCCOPY)
    If (lnt_Result = 0) Then GoTo lbl_Abort
  
  End If
  
  Exit Function
  
lbl_Abort:
  
  BlitFast = False
  
End Function

Public Function Blit(lng_DestDC As Long, int_Left As Integer, int_Top As Integer, int_Width As Integer, int_Height As Integer, int_Frame As Integer, Optional bln_BlitTransParent As Boolean = True) As Boolean
  
  Dim int_SrcLeft As Integer
  Dim int_SrcTop  As Integer
  Dim lnt_Result  As Long
  
  On Error GoTo lbl_Abort
  
  Blit = True
  
  int_SrcLeft = int_FrameWidth * (int_Frame Mod int_FramesWide)
  int_SrcTop = IIf(int_FrameWidth = 1, 0, int_FrameHeight * (int_Frame \ FramesWide))
  
  If bln_BlitTransParent Then
    
    lnt_Result = StretchBlt(lng_DestDC, int_Left, int_Top, int_Width, int_Height, lng_MaskDC, int_SrcLeft, int_SrcTop, int_FrameWidth, int_FrameHeight, SRCAND)
    If (lnt_Result = 0) Then GoTo lbl_Abort
    
    lnt_Result = StretchBlt(lng_DestDC, int_Left, int_Top, int_Width, int_Height, lng_InvertImageDC, int_SrcLeft, int_SrcTop, int_FrameWidth, int_FrameHeight, SRCPAINT)
    If (lnt_Result = 0) Then GoTo lbl_Abort
  
  Else
    
    lnt_Result = StretchBlt(lng_DestDC, int_Left, int_Top, int_Width, int_Height, lng_ImageDC, int_SrcLeft, int_SrcTop, int_FrameWidth, int_FrameHeight, SRCCOPY)
    If (lnt_Result = 0) Then GoTo lbl_Abort
  
  End If
  
  Exit Function
  
lbl_Abort:
  
  Blit = False
  
End Function

Public Function MakeBitmap(ByVal int_Width As Integer, ByVal int_Height As Integer, ByVal lng_ParentDC As Long) As Boolean
  
  Dim lng_Result  As Long
  Dim lng_Bitmap  As Long
  
  Dim bln_Return  As Boolean
  
  On Error GoTo lbl_Abort
   
  If bln_BitmapLoaded Then CleanUp
    
  bln_Return = True
  
  lng_Width = CLng(int_Width)
  lng_Height = CLng(int_Height)
  
  If (int_FramesWide = 0) Then int_FramesWide = 1
  If (int_FramesHigh = 0) Then int_FramesHigh = 1
  
  int_FrameCount = int_FramesWide * int_FramesHigh
  
  int_FrameWidth = lng_Width \ int_FramesWide
  int_FrameHeight = lng_Height \ int_FramesHigh
  
  lng_ImageDC = CreateCompatibleDC(lng_ParentDC)
  If (lng_ImageDC = 0) Then bln_Return = False
    
  lng_Bitmap = CreateCompatibleBitmap(lng_ParentDC, int_Width, int_Height)
  If (lng_Bitmap = 0) Then bln_Return = False
    
  lng_Result = SelectObject(lng_ImageDC, lng_Bitmap)
  If (lng_Result = 0) Then bln_Return = False
  
  lng_Result = BitBlt(lng_ImageDC, 0, 0, lng_Width, lng_Height, lng_ImageDC, 0, 0, SRCERASE)
  If (lng_Result = 0) Then bln_Return = False
  
  lng_Result = DeleteObject(lng_Bitmap)
  If (lng_Result = 0) Then bln_Return = False
    
  MakeBitmap = bln_Return
  bln_BitmapLoaded = bln_Return
  
  Exit Function
  
lbl_Abort:
  
  MakeBitmap = False
  bln_BitmapLoaded = False
  
End Function

Public Function LoadBitmapFromRes(var_Image As Variant, var_Type As Variant) As Boolean
  
  Dim byt_Array()   As Byte
  Dim bln_Return    As Boolean
  Dim int_FreeFile  As Integer
  
  On Error GoTo lbl_Abort
  
  bln_Return = True
  
  If bln_BitmapLoaded Then CleanUp
  
  Select Case var_Type
    Case vbResBitmap
      SavePicture LoadResPicture(var_Image, vbResBitmap), App.Path & "\temp.bmp"
    Case vbResIcon
      SavePicture LoadResPicture(var_Image, vbResIcon), App.Path & "\temp.bmp"
    Case vbResCursor
      SavePicture LoadResPicture(var_Image, vbResCursor), App.Path & "\temp.bmp"
    Case Else
    
      byt_Array = LoadResData(var_Image, var_Type)
      int_FreeFile = FreeFile
      
      Open App.Path & "\temp.bmp" For Binary Access Write As #int_FreeFile
        Put #int_FreeFile, , byt_Array
      Close #int_FreeFile
  
  End Select
  
  bln_Return = LoadBitmapFromFile(App.Path & "\temp.bmp")
  
  Kill App.Path & "\temp.bmp"
  
  LoadBitmapFromRes = bln_Return
  bln_BitmapLoaded = bln_Return

  Exit Function
  
lbl_Abort:

  LoadBitmapFromRes = False
  bln_BitmapLoaded = False

End Function

Public Function LoadBitmapFromFile(str_File As String) As Boolean
  
  Dim bmp_Data      As Bitmap

  Dim lng_Result    As Long
  Dim lng_Bitmap    As Long
  Dim bln_Return    As Boolean
  
  On Error GoTo lbl_Abort
    
  If bln_BitmapLoaded Then CleanUp

  bln_Return = True
  
  SavePicture LoadPicture(str_File), str_File
  
  lng_Bitmap = LoadImage(0, str_File, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE)
  If (lng_Bitmap = 0) Then bln_Return = False
    
  lng_Result = GetObject(lng_Bitmap, Len(bmp_Data), bmp_Data)
  If (lng_Result = 0) Then bln_Return = False
  
  lng_Width = bmp_Data.bmWidth
  lng_Height = bmp_Data.bmHeight
  
  If (int_FramesWide = 0) Then int_FramesWide = 1
  If (int_FramesHigh = 0) Then int_FramesHigh = 1
  
  int_FrameCount = int_FramesWide * int_FramesHigh
  
  int_FrameWidth = lng_Width \ int_FramesWide
  int_FrameHeight = lng_Height \ int_FramesHigh

  lng_ImageDC = CreateCompatibleDC(0)
  lng_InvertImageDC = CreateCompatibleDC(0)
  lng_MaskDC = CreateCompatibleDC(0)
  If (lng_ImageDC = 0) Or (lng_InvertImageDC = 0) Or (lng_MaskDC = 0) Then bln_Return = False
  
  lng_ImageOld = SelectObject(lng_ImageDC, lng_Bitmap)
  If (lng_Result = 0) Then bln_Return = False
  
  lng_InvertImage = CreateCompatibleBitmap(lng_ImageDC, lng_Width, lng_Height)
  If (lng_InvertImage = 0) Then bln_Return = False
  
  lng_InvertImageOld = SelectObject(lng_InvertImageDC, lng_InvertImage)
  If (lng_Result = 0) Then bln_Return = False
  
  lng_Mask = CreateCompatibleBitmap(lng_MaskDC, lng_Width, lng_Height)
  If (lng_Mask = 0) Then bln_Return = False
  
  lng_MaskOld = SelectObject(lng_MaskDC, lng_Mask)
  If (lng_Result = 0) Then bln_Return = False
  
  lng_Result = BitBlt(lng_InvertImageDC, 0, 0, lng_Width, lng_Height, lng_ImageDC, 0, 0, SRCCOPY)
  If (lng_Result = 0) Then bln_Return = False
  
  lng_Result = BitBlt(lng_MaskDC, 0, 0, lng_Width, lng_Height, lng_ImageDC, 0, 0, SRCCOPY)
  If (lng_Result = 0) Then bln_Return = False
  
  lng_Result = SetBkColor(lng_InvertImageDC, vbBlack)
  If (lng_Result = -1) Then bln_Return = False
  
  lng_Result = SetTextColor(lng_InvertImageDC, vbWhite)
  If (lng_Result = -1) Then bln_Return = False
  
  lng_Result = BitBlt(lng_InvertImageDC, 0, 0, lng_Width, lng_Height, lng_MaskDC, 0, 0, SRCAND)
  If (lng_Result = 0) Then bln_Return = False
    
  lng_Result = DeleteObject(lng_Bitmap)
  If (lng_Result = 0) Then bln_Return = False
  
  LoadBitmapFromFile = bln_Return
  bln_BitmapLoaded = bln_Return
  
  Exit Function
  
lbl_Abort:
  
  LoadBitmapFromFile = False
  bln_BitmapLoaded = False
  
End Function

Private Sub Class_Terminate()
  
  CleanUp

End Sub
  

