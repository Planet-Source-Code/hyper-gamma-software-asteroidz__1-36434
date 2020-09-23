VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_PHShot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements cls_PhysicsHandler

Private obj_Parent        As Object

Private pnt_Position      As typ_Point
Private pnt_Velocity      As typ_Point
Private pnt_Acceleration  As typ_Point

Private int_Angle         As Integer
Private int_SpeedLimit    As Integer
Private int_Left          As Integer
Private int_Top           As Integer
'

Private Property Get cls_PhysicsHandler_Parent() As Object
  
  Set cls_PhysicsHandler_Parent = obj_Parent
  
End Property

Private Property Set cls_PhysicsHandler_Parent(ByVal obj_NewValue As Object)
  
  Set obj_Parent = obj_NewValue
  
End Property

Private Property Get cls_PhysicsHandler_Angle() As Integer

  cls_PhysicsHandler_Angle = int_Angle
  
End Property

Private Property Let cls_PhysicsHandler_Angle(ByVal int_NewValue As Integer)
  
  int_Angle = IIf(int_NewValue < 0, 360 + int_NewValue, int_NewValue Mod 360)
  
  pnt_Velocity.X = cls_PhysicsHandler_Velocity * Cos(DegreeToRadian(360 - int_Angle))
  pnt_Velocity.Y = cls_PhysicsHandler_Velocity * Sin(DegreeToRadian(360 - int_Angle))

End Property

Private Property Get cls_PhysicsHandler_Velocity() As Single
  
  cls_PhysicsHandler_Velocity = Sqr(pnt_Velocity.X * pnt_Velocity.X + pnt_Velocity.Y * pnt_Velocity.Y)

End Property

Private Property Let cls_PhysicsHandler_Velocity(ByVal sng_NewValue As Single)
    
  pnt_Velocity.X = sng_NewValue * Cos(DegreeToRadian(360 - cls_PhysicsHandler_Angle))
  pnt_Velocity.Y = sng_NewValue * Sin(DegreeToRadian(360 - cls_PhysicsHandler_Angle))

End Property

Private Property Get cls_PhysicsHandler_SpeedLimit() As Integer
  
  cls_PhysicsHandler_SpeedLimit = int_SpeedLimit
  
End Property

Private Property Let cls_PhysicsHandler_SpeedLimit(ByVal int_NewValue As Integer)
  
  int_SpeedLimit = int_NewValue
  
End Property

Private Property Let cls_PhysicsHandler_Left(ByVal int_NewValue As Integer)
  
  pnt_Position.X = int_NewValue
  
End Property

Private Property Get cls_PhysicsHandler_Left() As Integer

  cls_PhysicsHandler_Left = int_Left
  
End Property

Private Property Let cls_PhysicsHandler_Top(ByVal int_NewValue As Integer)
  
  pnt_Position.Y = int_NewValue
    
End Property

Private Property Get cls_PhysicsHandler_Top() As Integer

  cls_PhysicsHandler_Top = pnt_Position.Y
  
End Property

Private Sub cls_PhysicsHandler_Accellerate(Optional sng_Power As Single = 1)
  
  If cls_PhysicsHandler_Velocity < int_SpeedLimit Then
    
    pnt_Acceleration.X = sng_Power * Cos(DegreeToRadian(360 - int_Angle))
    pnt_Acceleration.Y = sng_Power * Sin(DegreeToRadian(360 - int_Angle))
    
    pnt_Velocity.X = pnt_Velocity.X + pnt_Acceleration.X
    pnt_Velocity.Y = pnt_Velocity.Y + pnt_Acceleration.Y
      
  End If
  
End Sub

Private Sub cls_PhysicsHandler_Deccellerate(Optional sng_Power As Single = 0.99)
    
  If cls_PhysicsHandler_Velocity > 0.25 Then
    pnt_Velocity.X = pnt_Velocity.X * sng_Power
    pnt_Velocity.Y = pnt_Velocity.Y * sng_Power
  Else
    pnt_Velocity.X = 0
    pnt_Velocity.Y = 0
  End If
    
End Sub

Private Sub cls_PhysicsHandler_Update()
           
  pnt_Position.X = pnt_Position.X + pnt_Velocity.X
  pnt_Position.Y = pnt_Position.Y + pnt_Velocity.Y
    
  With cls_PhysicsHandler_Parent
        
    If pnt_Position.X > .Parent.Parent.Parent.ScaleWidth Then .Terminating = True
    If pnt_Position.X < -.Width Then .Terminating = True
    If pnt_Position.Y > .Parent.Parent.Parent.ScaleHeight Then .Terminating = True
    If pnt_Position.Y < -.Height Then .Terminating = True
    
    .Left = pnt_Position.X
    .Top = pnt_Position.Y
    
  End With
    
End Sub

