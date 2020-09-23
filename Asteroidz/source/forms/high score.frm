VERSION 5.00
Begin VB.Form frm_HighScores 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "High Scores"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3795
   Icon            =   "high score.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   168
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   253
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frm_HighScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  
  Private Labels      As cls_Labels
  Private bmp_Fontmap As cls_Bitmap
'

Public Property Let ScoreList(str_NewValue As String)
  
  With Labels.Item("SCORES")
    .Text = str_NewValue
    .Left = Labels.Item("NAMES").Left + Labels.Item("NAMES").Width + bmp_Fontmap.FrameWidth
    .Top = bmp_Fontmap.FrameHeight
  End With
  
End Property

Public Property Let NameList(str_NewValue As String)

  With Labels.Item("NAMES")
    .Text = str_NewValue
    .Left = bmp_Fontmap.FrameWidth
    .Top = bmp_Fontmap.FrameHeight
  End With
  
End Property

Public Sub Render()
  
  Labels.Render

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = vbKeyReturn Then Unload Me
  
End Sub

Private Sub Form_Load()

  Picture = Image
  
  BackColor = vbWhite
  
  Set bmp_Fontmap = New cls_Bitmap
  
  With bmp_Fontmap
    .FramesWide = 49
    .FramesHigh = 2
    .LoadBitmapFromRes "FONTMAP", "GRAPHICS"
  End With
  
  Set Labels = New cls_Labels
  
  With Labels
    
    Set .Bitmap = bmp_Fontmap
    
    .DestDC = hDC
    
    With .Add("NAMES")
      .Activate
      .Visible = True
    End With
    
    With .Add("SCORES")
      .Activate
      .Visible = True
    End With
  
  End With
  
  Refresh

End Sub

Private Sub Form_Unload(Cancel As Integer)

  Set Labels = Nothing
  Set bmp_Fontmap = Nothing
  
End Sub
