VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_Ship"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 
 
Public Parent               As cls_Main
Public Bitmap               As cls_Bitmap

Public Index                As Integer

Private obj_PhysicsHandler  As cls_PhysicsHandler
Private rct_Dest            As RECT

Private int_Angle           As Integer
Private str_Command         As String
Private bln_Accelerating    As Boolean
Private bln_Visible         As Boolean
Private int_Frame           As Integer
Private sng_Velocity        As Single
'

Property Get Width() As Integer
  
  Width = rct_Dest.Right - rct_Dest.Left

End Property

Property Get Height() As Integer
  
  Height = rct_Dest.Bottom - rct_Dest.Top

End Property

Public Property Get Left() As Integer
  
  Left = rct_Dest.Left
  
End Property

Public Property Let Left(ByVal int_NewValue As Integer)
    
  With rct_Dest
    .Left = int_NewValue
    .Right = .Left + Bitmap.FrameWidth
  End With

End Property

Public Property Get Top() As Integer
  
  Top = rct_Dest.Top
  
End Property

Public Property Let Top(ByVal int_NewValue As Integer)
    
  With rct_Dest
    .Top = int_NewValue
    .Bottom = .Top + Bitmap.FrameHeight
  End With

End Property

Public Function Hit(obj_Asteroid As cls_Asteroid) As Boolean
      
  Dim rct_Return  As RECT
  Dim rct_Rect    As RECT
  
  Dim bln_Return As Boolean
  
  If (Not obj_Asteroid.Ready) Then Exit Function
  
  With obj_Asteroid
    rct_Rect = GetRect(.Left, .Top, .Left + .Width, .Top + .Height)
  End With
                        
  bln_Return = IntersectRect(rct_Return, rct_Rect, rct_Dest)
  
  If bln_Return Then Parent.Shield = Parent.Shield - obj_Asteroid.Power
  
  Hit = bln_Return
  
End Function

Public Sub Activate()
  
  Set Bitmap = Parent.Bitmaps.Item("SHIP")
  
  With rct_Dest
  
    .Left = (Parent.Parent.ScaleWidth - Bitmap.FrameWidth) / 2
    .Top = (Parent.Parent.ScaleHeight - Bitmap.FrameHeight) / 2
    .Right = .Left + Bitmap.FrameWidth
    .Bottom = .Top + Bitmap.FrameHeight
        
  End With
  
  Set obj_PhysicsHandler = MakePH("SHIP")
  
  With obj_PhysicsHandler
    Set .Parent = Me
    .SpeedLimit = 20
    .Angle = 90
    .Left = rct_Dest.Left
    .Top = rct_Dest.Top
  End With
  
  bln_Visible = True
    
End Sub

Public Sub Update()
  
  Dim obj_Asteroid As cls_Asteroid
  
  If bln_Accelerating Then obj_PhysicsHandler.Accellerate
    
  If Not bln_Accelerating Then obj_PhysicsHandler.Deccellerate
  
  obj_PhysicsHandler.Update
       
  int_Frame = (obj_PhysicsHandler.Angle Mod 360) \ 10
  If bln_Accelerating Then int_Frame = int_Frame + 36

End Sub

Public Sub Render()
  
  If bln_Visible Then Bitmap.BlitFast Parent.Parent.hDC, CInt(rct_Dest.Left), CInt(rct_Dest.Top), int_Frame
  
End Sub

Public Property Get Command() As String
  
  Command = str_Command
  
End Property

Public Property Let Command(str_NewValue As String)
  
  Static bln_LeftGun As Boolean
  Dim int_FireAngle As Integer
  Dim str_Shot As String
  
  bln_Accelerating = False
  
  str_Command = str_NewValue
  
  Select Case str_Command
    Case "TURNRIGHT"
      obj_PhysicsHandler.Angle = obj_PhysicsHandler.Angle - 5
    Case "TURNLEFT"
      obj_PhysicsHandler.Angle = obj_PhysicsHandler.Angle + 5
    Case "ACCELERATE"
      bln_Accelerating = True
    Case Else
    
      If Left(str_Command, 4) = "FIRE" Then
                
        bln_LeftGun = Not bln_LeftGun
        
        str_Shot = IIf((str_Command = "FIREPULSE"), "IONPULSE", "SEEKER")
        
        With Parent.Shots.Add(str_Shot)
          
          int_FireAngle = (obj_PhysicsHandler.Angle \ 10) * 10
          
          If bln_LeftGun Then
            .Left = (rct_Dest.Left + Width / 2) + (22 * Cos(DegreeToRadian(int_FireAngle + 19)))
            .Top = (rct_Dest.Top + Height / 2) - (22 * Sin(DegreeToRadian(int_FireAngle + 19)))
          Else
            .Left = (rct_Dest.Left + Width / 2) + (22 * Cos(DegreeToRadian(int_FireAngle - 19)))
            .Top = (rct_Dest.Top + Height / 2) - (22 * Sin(DegreeToRadian(int_FireAngle - 19)))
          End If
          
          .Angle = (obj_PhysicsHandler.Angle * 10) \ 10
          
          .Activate
          
        End With
      
      End If
    
  End Select
  
End Property

Private Sub Class_Terminate()
  
  Set obj_PhysicsHandler = Nothing
  
End Sub
