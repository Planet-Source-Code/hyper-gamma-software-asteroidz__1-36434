VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_Shot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Bitmap               As cls_Bitmap

Public Terminating          As Boolean
Public Power                As Long
Public ExploStyle           As String
Public ShotType             As String

Private obj_PhysicsHandler  As cls_PhysicsHandler
Private obj_Parent          As cls_Shots
Private obj_Target          As cls_Asteroid
Private rct_Dest            As RECT

Private int_Angle           As Single
Private sng_Speed           As Single
Private sng_DeltaX          As Single
Private sng_DeltaY          As Single
Private bln_TargetFound     As Boolean
Private sng_Top             As Single
Private sng_Left            As Single
Private lng_StartFrame      As Long
'

Public Property Get Parent() As cls_Shots
  
  Set Parent = obj_Parent
  
End Property

Public Property Set Parent(cls_NewValue As cls_Shots)
  
  Set obj_Parent = cls_NewValue
  
  lng_StartFrame = obj_Parent.Parent.FrameCount

End Property

Public Property Get Speed() As Single
  
  Speed = sng_Speed
  
End Property

Public Property Let Speed(sng_NewValue As Single)
  
  sng_Speed = sng_NewValue
  
End Property

Property Get Width() As Integer
  
  Width = Bitmap.FrameWidth
  
End Property

Property Get Height() As Integer
  
  Height = Bitmap.FrameHeight
  
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

Property Get Angle() As Integer

  If ObjPtr(obj_PhysicsHandler) Then Angle = obj_PhysicsHandler.Angle
  
End Property

Property Let Angle(int_NewValue As Integer)

  int_Angle = int_NewValue
  
  If ObjPtr(obj_PhysicsHandler) Then obj_PhysicsHandler.Angle = int_Angle

End Property

Public Sub Activate()

  Set obj_PhysicsHandler = MakePH("SHOT")
  
  With obj_PhysicsHandler
    Set .Parent = Me
    .Angle = int_Angle
    .Velocity = sng_Speed
    .Left = rct_Dest.Left
    .Top = rct_Dest.Top
  End With

End Sub

Public Sub Update()
            
  Dim PointA As typ_Point
  Dim PointB As typ_Point
  
  Dim int_TargetAngle As Integer
  
  Select Case ShotType
  
    Case "IONPULSE"
      
      If obj_Parent.Parent.FrameCount > lng_StartFrame + 100 Then Terminating = True

      sng_Left = sng_Left + sng_DeltaX
      sng_Top = sng_Top + sng_DeltaY
          
    Case "SEEKER"
    
      If obj_Parent.Parent.FrameCount > lng_StartFrame + 1000 Then Terminating = True
      
      If Not bln_TargetFound Then Set obj_Target = obj_Parent.Parent.Asteroids.WithinRange(rct_Dest.Left + (rct_Dest.Right - rct_Dest.Left) / 2, rct_Dest.Top + (rct_Dest.Bottom - rct_Dest.Top) / 2, 150)
      
      If (Parent.Parent.FrameCount Mod 2 = 0) And (Not Terminating) Then
        With Parent.Parent.Effects.Add("SMOKE2")
          .Left = Left
          .Top = Top
          .Activate
        End With
      End If
      
      If ObjPtr(obj_Target) Then
              
        bln_TargetFound = True
        
        With PointA
          .X = rct_Dest.Left + (rct_Dest.Right - rct_Dest.Left) / 2
          .Y = rct_Dest.Top + (rct_Dest.Bottom - rct_Dest.Top) / 2
        End With
        
        With PointB
          .X = obj_Target.Left + (obj_Target.Width / 2)
          .Y = obj_Target.Top + (obj_Target.Height / 2)
        End With

        If GetDistance(PointA, PointB) > 300 Then bln_TargetFound = False
        If obj_Target.Terminating Then bln_TargetFound = False
        
        If bln_TargetFound Then
          
          int_TargetAngle = GetAngle(PointA, PointB)
          
          If Abs(int_TargetAngle - int_Angle) < 180 Then
            Angle = IIf(int_TargetAngle > Angle, Angle + 3, Angle - 3)
          Else
            Angle = IIf(int_TargetAngle < Angle, Angle + 3, Angle - 3)
          End If
        
        End If
      
      End If
    
  End Select

  obj_PhysicsHandler.Update
        
End Sub

Public Sub Render()
  
  Dim rct_Temp As RECT
  
  Dim sng_CurrentFrame As Single
  
  Select Case ShotType
  
    Case "SEEKER"
      
      sng_CurrentFrame = (Angle Mod 360) \ 10

    Case Else
    
      sng_CurrentFrame = IIf(sng_CurrentFrame < Bitmap.FrameCount - 1, sng_CurrentFrame + 1, 0)

  End Select
    
  With rct_Temp
    .Left = rct_Dest.Left - (Bitmap.FrameWidth \ 2)
    .Top = rct_Dest.Top - (Bitmap.FrameHeight \ 2)
    .Right = rct_Dest.Right
    .Bottom = rct_Dest.Bottom
  End With
  
  Bitmap.BlitFast Parent.Parent.Parent.hDC, CInt(rct_Temp.Left), CInt(rct_Temp.Top), CInt(sng_CurrentFrame)

End Sub
