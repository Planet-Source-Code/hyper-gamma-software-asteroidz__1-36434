VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_Asteroid"
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
  
  Private obj_Parent As cls_Asteroids
  Private obj_Bitmap As cls_Bitmap
  
  Public Shield As Integer
  Public Power As Single
  Public Angle As Integer
  Public Index As Long
  Public Speed As Single
  Public Dying As Boolean
  Public DeltaX As Single
  Public DeltaY As Single
  Public ShotByPlayer As Boolean
  Public Ready As Boolean
  
  Private obj_PhysicsHandler  As cls_PhysicsHandler
  Private rct_Dest As RECT
  
  Private lng_StartTime As Long
  Private bln_AnimAdvance As Boolean
  Private sng_AnimationDelay As Single
  Private str_Type As String
  Private bln_Terminating As Boolean
  Private sng_CurrentFrame As Single
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
    
    If ObjPtr(obj_PhysicsHandler) Then obj_PhysicsHandler.Left = .Left
  
  End With
  
End Property

Public Property Get Top() As Integer
  
  Top = rct_Dest.Top
  
End Property

Public Property Let Top(ByVal int_NewValue As Integer)
    
  With rct_Dest
  
    .Top = int_NewValue
    .Bottom = .Top + Bitmap.FrameHeight
    
    If ObjPtr(obj_PhysicsHandler) Then obj_PhysicsHandler.Top = .Top
  
  End With
    
End Property

Property Get Bitmap() As cls_Bitmap
  
  Set Bitmap = obj_Bitmap
  
End Property

Property Set Bitmap(obj_NewValue As cls_Bitmap)

  Set obj_Bitmap = obj_NewValue
  
End Property

Property Get EnemyType() As String
  
  EnemyType = str_Type
  
End Property

Property Let EnemyType(str_NewValue As String)

  str_Type = str_NewValue
  
End Property

Property Get Terminating() As Integer
  
  Terminating = bln_Terminating
  
End Property

Property Let Terminating(bln_NewValue As Integer)

  bln_Terminating = bln_NewValue
  
End Property

Property Get Parent() As cls_Asteroids
  
  Set Parent = obj_Parent
  
End Property

Property Set Parent(obj_NewValue As cls_Asteroids)

  Set obj_Parent = obj_NewValue
  
End Property

Public Sub Update()
  
  Dim tmp_item1 As cls_Shot
  Dim tmp_Item2 As cls_Asteroid
  
  Dim str_Effect As String
  
  obj_PhysicsHandler.Update
  
  With rct_Dest
  
    .Left = Left
    .Top = Top
    .Right = .Left + Bitmap.FrameWidth
    .Bottom = .Top + Bitmap.FrameHeight
    
    Set tmp_item1 = Parent.Parent.Shots.Hit(Me)
    
  End With
  
  If GetTickCount >= (lng_StartTime + 1000) Then Ready = True
  
  If ObjPtr(tmp_item1) Then
    Shield = Shield - tmp_item1.Power
    tmp_item1.Terminating = True
    If Shield <= 0 Then
      Dying = True
      ShotByPlayer = True
    Else
      
      str_Effect = IIf(tmp_item1.ShotType = "IONPULSE", "HIT2", "HIT1")
      
      With Parent.Parent.Effects.Add(str_Effect)
        Set .ParentObject = Me
        .Left = tmp_item1.Left
        .Top = tmp_item1.Top
        .Activate
      End With
    End If
  End If
  
  Set tmp_Item2 = Parent.Hit(Me)
  
  If ObjPtr(tmp_Item2) Then
    'tmp_Item2.Dying = True
    'Dying = True
  End If
  
  If Parent.Parent.Player.Hit(Me) Then Dying = True
  
  sng_CurrentFrame = IIf(sng_CurrentFrame < Bitmap.FrameCount - 1, sng_CurrentFrame + sng_AnimationDelay, 0)

End Sub

Public Sub Render()
  
  Bitmap.BlitFast Parent.Parent.Parent.hDC, CInt(rct_Dest.Left), CInt(rct_Dest.Top), CInt(sng_CurrentFrame)

End Sub

Public Sub Activate()
  
  With rct_Dest
    .Left = Left
    .Top = Top
    .Right = .Left + Bitmap.FrameWidth
    .Bottom = .Top + Bitmap.FrameHeight
  End With
  
  Set obj_PhysicsHandler = MakePH("ASTEROID")
  
  With obj_PhysicsHandler
    Set .Parent = Me
    .Angle = Angle
    .Velocity = Speed
    .Left = rct_Dest.Left
    .Top = rct_Dest.Top
  End With
  
  lng_StartTime = GetTickCount
  
  sng_CurrentFrame = Rnd * Bitmap.FrameCount
  
  sng_AnimationDelay = 0.5
  
  bln_AnimAdvance = Int(Rnd * 2)
    
End Sub

