VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_Level"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Number As Integer

Private bln_NextLevel As Boolean
Private bln_SkipWait As Boolean


Public Sub Restart()
  
  If (Not bln_RestartLevel) Then bln_RestartLevel = True
  
End Sub

Public Sub Update()
  
  Static lng_StartFrame As Long
  Static bln_Waiting    As Boolean
  
  Dim bln_SkipWait      As Boolean
  
  If (Parent.Asteroids.Count = 0) And (Parent.Effects.Count = 0) Then
                      
    If Not bln_NextLevel Then
      
      Parent.Message = "Level: " & Format((Number + 1), "00")
      bln_NextLevel = True
      
      If (Not bln_RestartLevel) Then Parent.InitScreen
      
      If Parent.Sound Then Parent.Sounds.Item("LEVEL").PlaySound
      
    End If
  
    If (Not obj_Message.Visible) Then
    
      Number = Number + 1
      
      Parent.Asteroids.Activate Number
      Parent.Player.Activate
      Parent.Shots.Clear
      
      bln_RestartLevel = False

      Shield = 100
      
      bln_NextLevel = False
      
    End If
        
  End If
  
End Sub

