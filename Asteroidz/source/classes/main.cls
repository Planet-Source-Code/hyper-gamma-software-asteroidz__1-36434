VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Parent               As frm_Main

Public Player               As cls_Ship
Public InputDevice          As cls_Input
Public Bitmaps              As cls_Bitmaps
Public Screen               As cls_Bitmap
Public Asteroids            As cls_Asteroids
Public Effects              As cls_Effects
Public Shots                As cls_Shots
Public Sounds               As cls_Sounds
Public Labels               As cls_Labels
Public Scores               As cls_Scores

Public Sound                As Boolean
Public Terminating          As Boolean
Public FrameCount           As Long
Public FrameRate            As Integer
Public Score                As Integer
Public Shield               As Integer
Public Lives                As Integer

Private obj_Screen          As cls_Res
Private obj_BmpStars        As cls_Bitmap
Private obj_Message         As cls_Label

Private lng_Score           As Long
Private lng_CumalativeScore As Long
Private bln_NextLevel       As Boolean
Private bln_RestartLevel    As Boolean
Private bln_Terminated      As Boolean
Private int_Level           As Integer
Private bln_Pause           As Boolean
Private bln_Trails          As Boolean
Private bln_Turbo           As Boolean
Private int_FrameLimit      As Integer
Private int_TargetFPS       As Integer
Private str_Message         As String
Private int_CharCounter     As Integer
Private lng_StartTime       As Long
Private lng_InitTime        As Long
Private bln_GameInit        As Boolean
'

Property Get Message() As String

  Message = obj_Message.Text
  
End Property

Property Let Message(str_NewValue As String)
    
  If (str_NewValue <> str_Message) Then
  
    int_CharCounter = 0
    str_Message = str_NewValue
    
    With obj_Message
      .Text = str_NewValue & "_"
      .Left = (Parent.ScaleWidth - .Width) / 2
      .Top = (Parent.ScaleHeight - .Height) / 2
      .Text = ""
      .Visible = True
    End With
    
    lng_StartTime = GetTickCount()

  End If
  
End Property
  
Private Sub InitBitmaps()
  
  Dim int_Counter As Integer
  
  Set Bitmaps = New cls_Bitmaps
  
  With Bitmaps
    
    With .Add("SHIP")
      .FramesWide = 36
      .FramesHigh = 2
      .LoadBitmapFromRes "SHIP", "GRAPHICS"
    End With
    
    For int_Counter = 1 To 4
    
      With .Add("ASTEROID" & int_Counter)
        .FramesWide = 6
        .FramesHigh = 6
        .LoadBitmapFromRes "ASTEROID" & int_Counter, "GRAPHICS"
      End With
      
      If int_Counter < 3 Then
      
        With .Add("EXPLOSION" & int_Counter)
          .FramesWide = 6
          .FramesHigh = 5
          .LoadBitmapFromRes "EXPLOSION" & int_Counter, "GRAPHICS"
        End With
        
        With .Add("HIT" & int_Counter)
          .FramesWide = 3
          .FramesHigh = 3
          .LoadBitmapFromRes "HIT" & int_Counter, "GRAPHICS"
        End With
        
        With .Add("SMOKE" & int_Counter)
          .FramesWide = 4
          .FramesHigh = 2
          .LoadBitmapFromRes "SMOKE" & int_Counter, "GRAPHICS"
        End With
      
      End If
      
    Next
        
    With .Add("IONPULSE")
      .FramesWide = 4
      .FramesHigh = 2
      .LoadBitmapFromRes "IONPULSE", "GRAPHICS"
    End With
    
    With .Add("SEEKER")
      .FramesWide = 6
      .FramesHigh = 6
      .LoadBitmapFromRes "SEEKER", "GRAPHICS"
    End With
    
    With .Add("STAR")
      .FramesWide = 5
      .FramesHigh = 1
      .LoadBitmapFromRes "STAR", "GRAPHICS"
    End With
    
    With .Add("PLANETS")
      .FramesWide = 3
      .FramesHigh = 3
      .LoadBitmapFromRes "PLANETS", "GRAPHICS"
    End With
            
    With .Add("FONTMAP")
      .FramesWide = 49
      .FramesHigh = 2
      .LoadBitmapFromRes "FONTMAP", "GRAPHICS"
    End With
    
    With .Add("SHIELD")
      .FramesWide = 20
      .FramesHigh = 1
      .LoadBitmapFromRes "SHIELD", "GRAPHICS"
    End With
    
    .Add("SCREEN").MakeBitmap Parent.ScaleWidth, Parent.ScaleHeight, Parent.hDC
      
  End With
  
  Set Screen = Bitmaps.Item("SCREEN")
  
End Sub

Private Sub InitSounds()
  
  Set Sounds = New cls_Sounds
  
  With Sounds
    .Add("SHOOT").LoadSoundFromRes "SHOOT", "SOUNDS"
    .Add("HIT").LoadSoundFromRes "HIT", "SOUNDS"
    .Add("EXPLOSION").LoadSoundFromRes "EXPLOSION", "SOUNDS"
    .Add("LEVEL").LoadSoundFromRes "LEVEL", "SOUNDS"
  End With
  
End Sub

Private Sub InitInput()

  Set InputDevice = New cls_Input
  Dim keys(11) As KeyCodeConstants
  
  With InputDevice
  
    .AddKey vbKeyF1
    .AddKey vbKeyF2
    .AddKey vbKeyF3
    .AddKey vbKeyF4
    .AddKey vbKeyF5
    .AddKey vbKeyF6
    
    .AddKey vbKeyControl
    .AddKey vbKeyEscape
    
    .AddKey vbKeyLeft
    .AddKey vbKeyRight
    .AddKey vbKeyUp
    .AddKey vbKeyDown
    
    .AddKey vbKeySpace
        
  End With
  
End Sub

Private Sub InitLabels()

  Set Labels = New cls_Labels
  
  With Labels
    
    Set .Parent = Me
    
    Set .Bitmap = Bitmaps.Item("FONTMAP")
    
    .DestDC = Parent.hDC
    
    With .Add("FPSCOUNTER")
      .Activate
      .Text = "FPS:    "
      .Left = (Parent.ScaleWidth - .Width) / 2
      .Top = Bitmaps.Item("FONTMAP").FrameWidth
    End With
    
    With .Add("MESSAGE")
      .Activate
      .Text = "Level:   "
      .Left = (Parent.ScaleWidth - .Width) / 2
      .Top = (Parent.ScaleHeight - .Height) / 2
    End With
    
    With .Add("SCORE")
      .Activate
      .Text = "Score: 0000000"
      .Left = Bitmaps.Item("FONTMAP").FrameWidth
      .Top = Parent.ScaleHeight - Bitmaps.Item("FONTMAP").FrameWidth * 2
      .Visible = True
    End With
    
    With .Add("SHIELD")
      .Activate
      .Text = "Shield: "
      .Left = Parent.ScaleWidth - Bitmaps.Item("SHIELD").Width - Bitmaps.Item("FONTMAP").FrameWidth - .Width
      .Top = Parent.ScaleHeight - Bitmaps.Item("FONTMAP").FrameWidth * 2
      .Visible = True
    End With
    
    With .Add("LIVES")
      .Activate
      .Text = " Lives: "
      .Left = Parent.ScaleWidth - Bitmaps.Item("SHIELD").Width - Bitmaps.Item("FONTMAP").FrameWidth - Labels.Item("SHIELD").Width
      .Top = Parent.ScaleHeight - Bitmaps.Item("FONTMAP").FrameWidth * 4
      .Visible = True
    End With
    
    Set obj_Message = .Item("MESSAGE")
    
  End With
  
End Sub

Private Sub InitScreen()

  Dim int_StarCounter As Integer
  
  Randomize
  
  Screen.Cls
  
  For int_StarCounter = 0 To 100
    Bitmaps.Item("STAR").BlitFast Screen.ImageDC, Rnd * Parent.ScaleWidth, Rnd * Parent.ScaleHeight, 5 - Int(Rnd * 5), False
  Next
  
  Bitmaps.Item("PLANETS").BlitFast Screen.ImageDC, Rnd * (Parent.ScaleWidth - Bitmaps.Item("PLANETS").FrameWidth), Rnd * (Parent.ScaleHeight - Bitmaps.Item("PLANETS").FrameHeight), int_Level Mod 9

End Sub

Public Sub Activate()
    
  FrameCount = 1
  
  InitBitmaps
  InitSounds
  InitInput
  InitLabels
  
  Set Player = New cls_Ship
  Set Player.Parent = Me
  
  Set Asteroids = New cls_Asteroids
  Set Asteroids.Parent = Me
  
  Set Shots = New cls_Shots
  Set Shots.Parent = Me
  
  Set Effects = New cls_Effects
  Set Effects.Parent = Me
  
  Set Scores = New cls_Scores
  
  Set obj_Screen = New cls_Res
  
  obj_Screen.ChangeRes 640, 480, 16
  
  lng_InitTime = GetTickCount()
    
  With GetConfig()
    Labels.Item("FPSCOUNTER").Visible = .FPS
    Sound = .Sound
    bln_Trails = .Trails
    bln_Turbo = .Turbo
  End With
  
  int_TargetFPS = 50
  int_FrameLimit = 1

  Shield = 100
  Lives = 3
  
  bln_GameInit = True
  
End Sub

Public Sub Execute()
  
  Dim tmp_Config As typ_GameConfig
  
  Dim lng_LastTick          As Long
  Dim lng_StartTime         As Long
  Dim lng_FrameRateCounter  As Long
  Dim str_Message           As String
    
  Do
  
    If Parent.WindowState <> vbMinimized Then
      
      If ((GetTickCount - lng_LastTick) >= int_FrameLimit) Then
                
        FrameCount = FrameCount + 1
        
        If GetTickCount > lng_StartTime + 1000 Then
          FrameRate = lng_FrameRateCounter
          Labels.Item("FPSCOUNTER").Text = "FPS: " & Format(FrameRate, "00")
          int_FrameLimit = IIf(FrameRate > int_TargetFPS, int_FrameLimit + 1, int_FrameLimit - 1)
          lng_FrameRateCounter = 0
          lng_StartTime = GetTickCount()
        End If
      
        lng_FrameRateCounter = lng_FrameRateCounter + 1
                        
        If Terminating Then
          
          lng_Score = lng_Score + Score
          If lng_Score < 0 Then lng_Score = 0
          Score = 0
          
          Message = "Game Over"
          
          If obj_Message.Text = "Game Over_" Then Exit Do
        
        End If
        
        Labels.Item("LIVES").Text = " Lives: " & Format(Lives, "00")
        Labels.Item("SCORE").Text = "Score: " & Format(lng_Score, "000000")
        
        Update
        Render
        
        lng_LastTick = GetTickCount
        
      End If
                                
    End If
    
    If Not (FrameCount Mod 5) Then DoEvents

  Loop
  
  With tmp_Config
    .FPS = Labels.Item("FPSCOUNTER").Visible
    .Sound = Sound
    .Trails = bln_Trails
    .Turbo = bln_Turbo
  End With
  
  SetConfig tmp_Config

  Scores.CheckHighScore lng_Score
  
End Sub

Private Sub DoLevels()
  
  Static lng_StartFrame As Long
  Static bln_Waiting    As Boolean
  
  Dim bln_SkipWait      As Boolean
  
  If (Asteroids.Count = 0) And (Effects.Count = 0) Then
                      
    If Not bln_NextLevel Then
      
      str_Message = ""
      Message = "Level: " & Format((int_Level + 1), "00")
      bln_NextLevel = True
      
      If (Not bln_RestartLevel) Then InitScreen
      
      If Sound Then Sounds.Item("LEVEL").PlaySound
      
    End If
  
    If (Not obj_Message.Visible) Then
    
      int_Level = int_Level + 1
      
      Asteroids.Activate int_Level
      Player.Activate
      Shots.Clear
      
      bln_RestartLevel = False

      Shield = 100
      
      bln_NextLevel = False
      
    End If
        
  End If
  
End Sub

Private Sub DoShield()
  
  Dim int_Counter     As Integer
  Dim int_ShieldLeft  As Integer
  Dim int_Left        As Integer
  Dim int_Top         As Integer
  
  Static bln_Visible  As Boolean
  Static int_Interval As Integer
  
  If (int_Interval = 0) Then int_Interval = 1
  
  int_ShieldLeft = Parent.ScaleWidth - Bitmaps.Item("FONTMAP").FrameWidth - Bitmaps.Item("SHIELD").Width
  int_Top = Parent.ScaleHeight - Bitmaps.Item("FONTMAP").FrameWidth * 2
    
  If (Shield <= 35) And (Shield >= 0) Then
    
    bln_Visible = (GetTickCount Mod 500 > 250)
    
    If (Not bln_RestartLevel) And (Not bln_NextLevel) Then
    
      If (FrameCount Mod int_Interval = 0) Then
        
        With Effects.Add("SMOKE" & CStr(Int(Rnd * 2) + 1))
          .Left = Player.Left + (Rnd * (Player.Width - Bitmaps.Item("SMOKE1").FrameWidth))
          .Top = Player.Top + (Rnd * (Player.Height - Bitmaps.Item("SMOKE1").FrameHeight))
          .Activate
        End With
        
        int_Interval = Int(Rnd * 30) + 1
        
      End If
    
    End If
    
  Else
    bln_Visible = True
  End If

  For int_Counter = 0 To (Shield / 5) - 1
    
    int_Left = int_ShieldLeft + int_Counter * Bitmaps.Item("SHIELD").FrameWidth
    
    If bln_Visible Then Bitmaps.Item("SHIELD").BlitFast Parent.hDC, int_Left, int_Top, int_Counter
  
  Next

End Sub

Public Sub Update()
  
  Dim bln_Underscore As Boolean
  Static lng_FireTime  As Long
  
  DoLevels
  
  If (Not Terminating) Then
  
    If Score > 0 Then
      Score = Score - 25
      lng_Score = lng_Score + 25
      lng_CumalativeScore = lng_CumalativeScore + 25
    ElseIf Score < 0 Then
      Score = Score + 25
      lng_Score = lng_Score - 25
    End If
    
    If lng_Score > 999999 Then lng_Score = 999999
    
    If lng_Score < 0 Then lng_Score = 0
    
    If lng_CumalativeScore >= 50000 Then
      Lives = Lives + 1
      lng_CumalativeScore = 0
    End If
    
  End If

  InputDevice.Refresh
  
  If InputDevice.KeyDown(vbKeyF1) And (Not bln_NextLevel) Then
    bln_Pause = Not bln_Pause
    Message = IIf(bln_Pause, "Paused", "Unpaused")
  End If

  If (Not bln_NextLevel) And (Not bln_Pause) Then
  
    With InputDevice
                    
      If .KeyUp(vbKeyEscape) Then Terminating = True
      
      Player.Command = ""
          
      If .KeyPressed(vbKeyUp) Then Player.Command = "ACCELERATE"
      If .KeyPressed(vbKeyRight) Then Player.Command = "TURNRIGHT"
      If .KeyPressed(vbKeyLeft) Then Player.Command = "TURNLEFT"
      If .KeyDown(vbKeySpace) Then Player.Command = "FIREPULSE"
      
      If .KeyDown(vbKeyControl) And GetTickCount > lng_FireTime + 1000 Then
        Player.Command = "FIRESEEKER"
        lng_FireTime = GetTickCount()
      End If
      
      If .KeyDown(vbKeyF2) Then
        Labels.Item("FPSCOUNTER").Visible = Not Labels.Item("FPSCOUNTER").Visible
        Message = IIf(Labels.Item("FPSCOUNTER").Visible, "FPS Counter On", "FPS Counter Off")
      End If
      
      If .KeyDown(vbKeyF3) Then
        bln_Trails = Not bln_Trails
        Message = IIf(bln_Trails, "Trails On", "Trails Off")
      End If
      
      If .KeyDown(vbKeyF4) Then
        Sound = Not Sound
        Message = IIf(Sound, "Sound On", "Sound Off")
      End If
      
      If .KeyDown(vbKeyF5) Then
        bln_Turbo = Not bln_Turbo
        int_TargetFPS = IIf(bln_Turbo, 500, 50)
        Message = IIf(bln_Turbo, "Turbo On", "Turbo Off")
      End If
            
      If Shield <= 0 And (Not bln_RestartLevel) Then
        
        If Lives = 1 Then Terminating = True
        
        bln_RestartLevel = True
        Score = Score - 5000
        
        With Effects.Add("EXPLOSION1")
          .Left = Player.Left + (Player.Width / 2)
          .Top = Player.Top + (Player.Height / 2)
          .Activate
        End With
        
        Asteroids.Clear
        
        Lives = Lives - 1
        int_Level = int_Level - 1
        
      End If
      
    End With
        
    If Not Terminating Then
      Shots.Update
      Asteroids.Update
      Player.Update
      Effects.Update
    End If
    
  End If
  
  If obj_Message.Visible Then If GetTickCount() >= (lng_StartTime + 3000) Then obj_Message.Visible = False
  
  If str_Message <> "" Then
    
    bln_Underscore = (GetTickCount Mod 1000 > 500)
        
    int_CharCounter = int_CharCounter + 1
    If int_CharCounter > Len(str_Message) Then int_CharCounter = Len(str_Message)
    
    obj_Message.Text = IIf(bln_Underscore, Left(str_Message, int_CharCounter) & "_", Left(str_Message, int_CharCounter))
  
  End If
  
End Sub

Public Sub Render()
  
  If (Not bln_Trails) Then Screen.BlitFast Parent.hDC, 0, 0, 0, False
  
  If Not bln_NextLevel Then
    Shots.Render
    If (Not bln_RestartLevel) Then Player.Render
    Asteroids.Render
    Effects.Render
  End If

  If (Not bln_NextLevel) Then DoShield
  
  Labels.Render
  
  Parent.Refresh
  
End Sub

Public Sub Terminate()

  Set obj_Screen = Nothing
  Set Player = Nothing
  Set InputDevice = Nothing
  Set Bitmaps = Nothing
  Set Asteroids = Nothing
  Set Shots = Nothing
  Set Effects = Nothing
  Set Sounds = Nothing
  Set Labels = Nothing
  Set Scores = Nothing
    
  bln_Terminated = True
  
End Sub

Private Sub Class_Terminate()

  If (Not bln_Terminated) Then Terminate

End Sub
