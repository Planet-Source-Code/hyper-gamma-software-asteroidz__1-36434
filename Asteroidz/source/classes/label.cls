VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_Label"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'Project Info...
' - Author:          Zach "Orion" Collette
' - Company:         Hyper-Gamma Software
' - Contact:         kingtheoden17@hotmail.com
' - Description:     label class
' - Test Platform:   Windows 98 SE
' - Processor:       P3 450MHz.
'------------------------------------------------------------
  
  Option Explicit
  
  'Property Variables
  Private obj_Parent  As cls_Labels
  Private obj_Bitmap As cls_Bitmap
  
  'Private Variables
  Private str_Text()      As String
  Private bln_Visible     As Boolean
  Private lng_Left        As Long
  Private lng_Top         As Long
  Private int_Width       As Integer
  Private int_Height      As Long
  Private int_CharWidth   As Integer
  Private int_CharHeight  As Integer
  '
  
Property Get Parent() As cls_Labels
  
  Set Parent = obj_Parent
  
End Property

Property Set Parent(obj_NewValue As cls_Labels)

  Set obj_Parent = obj_NewValue
  
End Property

Property Get Surface() As cls_Bitmap
  
  Set Surface = obj_Bitmap
  
End Property

Property Set Surface(obj_NewValue As cls_Bitmap)

  Set obj_Bitmap = obj_NewValue
  
End Property

Property Get Text() As String
  
  Text = Join(str_Text, vbNewLine)
  
End Property

Property Let Text(str_NewValue As String)
 
  Dim int_LineCounter As Integer
  Dim int_LineWidth   As Integer
  Dim int_Temp        As String

  str_Text = Split(str_NewValue, vbNewLine)
  
  For int_LineCounter = 0 To UBound(str_Text)
    int_Temp = Len(str_Text(int_LineCounter))
    If int_Temp > int_LineWidth Then int_LineWidth = int_Temp
  Next
  
  int_Width = int_LineWidth * int_CharWidth
  int_Height = (UBound(str_Text) + 1) * int_CharHeight

End Property

Property Get Visible() As Boolean
  
  Visible = bln_Visible
  
End Property

Property Let Visible(bln_NewValue As Boolean)

  bln_Visible = bln_NewValue
  
End Property

Property Get Left() As Long
  
  Left = lng_Left
  
End Property

Property Let Left(int_NewValue As Long)

  lng_Left = int_NewValue
  
End Property

Property Get Top() As Long
  
  Top = lng_Top
  
End Property

Property Let Top(int_NewValue As Long)

  lng_Top = int_NewValue
  
End Property

Property Get Width() As Integer

  Width = int_Width
  
End Property

Property Get Height() As Integer
  
  Height = int_Height
  
End Property

Public Sub Activate()
  
  Set obj_Bitmap = Parent.Bitmap
  
  int_CharWidth = obj_Bitmap.FrameWidth
  int_CharHeight = obj_Bitmap.FrameHeight
  
End Sub

Public Sub Render()
  
  Dim int_CharCounter As Integer
  Dim int_LineCounter As Integer
  
  Dim str_Char As String
  Dim int_NumLines As Integer
  Dim int_LineWidth As Integer
  Dim int_Frame As Integer
  Dim rct_Temp As RECT
  
  If Not bln_Visible Then Exit Sub
  
  For int_LineCounter = 0 To UBound(str_Text)
  
    int_LineWidth = Len(str_Text(int_LineCounter))
    
    For int_CharCounter = 1 To int_LineWidth
              
      str_Char = Mid$(str_Text(int_LineCounter), int_CharCounter, 1)
  
      int_Frame = (Asc(str_Char) - 32)
      
      If str_Char = "°" Then int_Frame = 94
      If str_Char = "±" Then int_Frame = 95
                
      With rct_Temp
        .Left = lng_Left + (int_CharCounter - 1) * int_CharWidth
        .Top = lng_Top + int_LineCounter * int_CharHeight
        .Right = .Left + int_Width
        .Bottom = .Top + int_Height
      End With
            
      obj_Bitmap.BlitFast Parent.DestDC, CInt(rct_Temp.Left), CInt(rct_Temp.Top), int_Frame
      
      int_LineWidth = int_LineWidth + 1
      
    Next
    
  Next
  
End Sub
