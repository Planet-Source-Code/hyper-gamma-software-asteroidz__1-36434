VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_Button"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Parent   As cls_Buttons

Public Enabled  As Boolean
Public Command  As String
Public Left     As Integer
Public Top      As Integer
Public Frame    As Integer

Private rct_Dest  As RECT
'

Public Function MouseOver(int_X As Integer, int_Y As Integer) As Boolean
  
  Dim rct_Mouse As RECT
  
  With rct_Mouse
    .Left = int_X
    .Top = int_Y
    .Right = .Left + 1
    .Bottom = .Top + 1
  End With
  
  If Collide(rct_Mouse, rct_Dest) Then MouseOver = True
  
End Function

Public Sub Render()
  
  Dim int_Frame As Integer
  
  int_Frame = IIf(Enabled, Frame + 4, Frame)
  
  Parent.Bitmap.BlitFast Parent.Parent.hDC, Left, Top, int_Frame
  
End Sub

Public Sub Activate()
  
  With rct_Dest
    .Left = Left
    .Top = Top
    .Right = Left + Parent.Bitmap.FrameWidth
    .Bottom = Top + Parent.Bitmap.FrameHeight
  End With
    
End Sub
