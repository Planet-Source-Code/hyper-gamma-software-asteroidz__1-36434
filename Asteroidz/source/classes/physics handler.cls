VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_PhysicsHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Parent     As Object

Public Left       As Integer
Public Top        As Integer
Public Angle      As Integer
Public SpeedLimit As Integer
Public Velocity   As Single

Public Sub Accellerate(Optional sng_Power As Single = 1)
End Sub

Public Sub Deccellerate(Optional sng_Power As Single = 0.99)
End Sub

Public Sub Update()
End Sub
