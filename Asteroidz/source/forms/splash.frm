VERSION 5.00
Begin VB.Form frm_Splash 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6150
   ControlBox      =   0   'False
   Icon            =   "splash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   291
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frm_Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public MouseX           As Integer
Public MouseY           As Integer
Public MousePressed     As Boolean

Private obj_Splash      As cls_Bitmap
Private obj_BmpButtons  As cls_Bitmap
Private obj_Scores      As cls_Scores
Private obj_Buttons     As cls_Buttons

Private bln_Quit        As Boolean
'

Private Sub Form_Load()
    
  Set obj_Splash = New cls_Bitmap
  
  With obj_Splash
    .FramesWide = 1
    .FramesHigh = 1
    .LoadBitmapFromRes "SPLASH", "GRAPHICS"
  End With
  
  Set obj_BmpButtons = New cls_Bitmap
  
  With obj_BmpButtons
    .FramesWide = 3
    .FramesHigh = 3
    .LoadBitmapFromRes "BUTTONS", "GRAPHICS"
  End With
  
  Set obj_Scores = New cls_Scores
  
  Set obj_Buttons = New cls_Buttons
  
  With obj_Buttons
    Set .Parent = Me
    Set .Bitmap = obj_BmpButtons
    .Activate
  End With
  
  Show
          
  Do
  
    obj_Splash.BlitFast hDC, 0, 0, 0, False
    
    obj_Buttons.Update
    obj_Buttons.Render
    
    DoEvents
    Refresh

  Loop Until bln_Quit
  
  Unload Me
  
End Sub

Public Sub Command(str_Name As String)
  
  MousePressed = False
  
  Select Case str_Name
  
    Case "PLAYGAME"
      
      Load frm_Main
      frm_Main.Show
            
    Case "VIEWSCORES"
      
      obj_Scores.ShowHighScores
      
    Case "VIEWREADME"
      
      OpenDoc AppPath & "readme.txt"
      
    Case "QUIT"
    
      bln_Quit = True
          
  End Select
  
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  MousePressed = True
  
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  MouseX = CInt(X)
  MouseY = CInt(Y)
  
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  MousePressed = False
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  Set obj_Splash = Nothing
  Set obj_Scores = Nothing
  Set obj_Buttons = Nothing
  Set obj_BmpButtons = Nothing
  
End Sub
