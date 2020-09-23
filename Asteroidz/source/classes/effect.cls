VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_Effect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Project Info...
' - Author:          Zach "Orion" Collette
' - Company:         Hyper-Gamma Software
' - Contact:         kingtheoden17@hotmail.com
' - Project Name:    Stinger
' - Description:     Arcade style shooter
' - Test Platform:   Windows 98 SE
' - Processor:       P3 450MHz.
'------------------------------------------------------------
  
  Option Explicit

  Public Parent As cls_Effects
  Public Bitmap As cls_Bitmap
  Public ParentObject As cls_Asteroid

  Public Angle As Integer
  Public Speed As Integer
  Public Phase As Long
  Public Terminating As Boolean
  Public EffectType As String
  Public Left As Single, Top As Single
  Public DeltaX As Single, DeltaY As Single
    
  Private rct_Dest As RECT
  Private sng_CurrentFrame As Single
  '
  
Public Sub Activate()
    
  With rct_Dest
    .Left = Left
    .Top = Top
    .Right = .Left + Bitmap.FrameWidth
    .Bottom = .Top + Bitmap.FrameHeight
  End With
  
  DeltaX = Cos(DegreeToRadian(360 - Angle)) * Speed
  DeltaY = Sin(DegreeToRadian(360 - Angle)) * Speed
      
End Sub

Public Sub Update()
  
  With rct_Dest
  
    If ParentObject Is Nothing Then
      Left = Left + DeltaX
      Top = Top + DeltaY
    Else
      Left = Left + ParentObject.DeltaX
      Top = Top + ParentObject.DeltaY
    End If
    
    .Left = Left
    .Top = Top
    .Right = .Left + Bitmap.FrameWidth
    .Bottom = .Top + Bitmap.FrameHeight
    
  End With
  
  If Int(sng_CurrentFrame) < Bitmap.FrameCount - 1 Then
    sng_CurrentFrame = sng_CurrentFrame + 1
  Else
    Terminating = True
  End If
  
End Sub

Public Sub Render()
  
  Dim rct_Temp As RECT
       
  With rct_Temp
    .Left = rct_Dest.Left - (Bitmap.FrameWidth \ 2)
    .Top = rct_Dest.Top - (Bitmap.FrameHeight \ 2)
    .Right = rct_Dest.Right
    .Bottom = rct_Dest.Bottom
  End With

  Bitmap.BlitFast Parent.Parent.Parent.hDC, CInt(rct_Temp.Left), CInt(rct_Temp.Top), CInt(sng_CurrentFrame)
  
End Sub
