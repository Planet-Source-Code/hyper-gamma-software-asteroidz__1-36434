VERSION 5.00
Begin VB.Form frm_Main 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9600
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "main.frx":0BC2
   MousePointer    =   99  'Custom
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Game As cls_Main

Private Sub Form_Load()
        
  BackColor = vbWhite
  
  Set Game = New cls_Main
  
  With Game
  
    Set .Parent = Me
    
    .Activate
    
    Show

    .Execute
    .Terminate
    
    Hide
    
  End With
  
  MousePointer = vbDefault
  
  Unload Me
        
End Sub

Private Sub Form_Paint()

  If ObjPtr(Game) Then Game.Render
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  Dim frm_Temp As Form
  
  Set Game = Nothing
  
  For Each frm_Temp In Forms
    Unload frm_Temp
  Next
  
  End

End Sub
