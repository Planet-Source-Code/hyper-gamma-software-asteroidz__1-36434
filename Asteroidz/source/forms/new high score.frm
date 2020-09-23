VERSION 5.00
Begin VB.Form frm_NewHighScore 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New High Score"
   ClientHeight    =   1050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4125
   Icon            =   "new high score.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   70
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   275
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "frm_NewHighScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  

  Private Labels      As cls_Labels
  Private bmp_Buffer  As cls_Bitmap
  Private bmp_Fontmap As cls_Bitmap
  Private snd_Error   As cls_Sound
  
  Private str_Name    As String
  '
  
Private Sub Form_Load()
  
  Picture = Image
  
  BackColor = vbWhite
    
  Set bmp_Buffer = New cls_Bitmap
  
  bmp_Buffer.MakeBitmap ScaleWidth, ScaleHeight, hDC
  
  Set bmp_Fontmap = frm_Main.Game.Bitmaps.Item("FONTMAP")
  
  Set Labels = New cls_Labels
  
  With Labels
    
    Set .Bitmap = bmp_Fontmap
    
    .DestDC = hDC
    
    With .Add("MESSAGE")
      .Activate
      .Text = "Please enter your Name."
      .Left = bmp_Fontmap.FrameWidth
      .Top = bmp_Fontmap.FrameHeight
      .Visible = True
    End With
        
    With .Add("INPUT")
      .Activate
      .Text = "           "
      .Left = (ScaleWidth - .Width) / 2
      .Top = bmp_Fontmap.FrameHeight * 3
      .Visible = True
    End With
  
  End With
  
  Set snd_Error = New cls_Sound
  
  snd_Error.LoadSoundFromRes "ERROR", "SOUNDS"
  
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
            
  Select Case KeyAscii
                            
    Case Asc("A") To Asc("z")
      str_Name = str_Name & Chr(KeyAscii)
    
    Case Asc("0") To Asc("9")
      str_Name = str_Name & Chr(KeyAscii)
    
    Case Asc("/")
      str_Name = str_Name & Chr(KeyAscii)
    
    Case Asc("'")
      str_Name = str_Name & Chr(KeyAscii)
    
    Case vbKeySpace
      str_Name = str_Name & Chr(KeyAscii)
      
    Case vbKeyBack

      If Len(str_Name) > 0 Then
      
        str_Name = Left(str_Name, Len(str_Name) - 1)
      
      Else
                
        snd_Error.PlaySound
                
      End If
    
    Case vbKeyReturn
      
      Unload Me
      
    Case Else
    
  End Select
      
  If Len(str_Name) > 10 Then
                            
    str_Name = Left(str_Name, 10)
    
    snd_Error.PlaySound
    
  End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  Set Labels = Nothing
  Set bmp_Buffer = Nothing
    
  frm_Main.Game.Scores.NewName = str_Name
  
End Sub

Private Sub Timer1_Timer()
  
  Static bln_Underscore As Boolean
  
  bln_Underscore = (GetTickCount Mod 1000) > 500

  bmp_Buffer.BlitFast hDC, 0, 0, 0, False
  
  Labels.Item("INPUT").Text = IIf(bln_Underscore, str_Name & "_", str_Name)
  Labels.Render
  
  Refresh
  
End Sub
