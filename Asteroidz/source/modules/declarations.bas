Attribute VB_Name = "mod_Declarations"
Option Explicit

Public Type typ_Point
  
  X As Single
  Y As Single

End Type

Public Type RECT

  Left    As Long
  Top     As Long
  Right   As Long
  Bottom  As Long
  
End Type

Type typ_HighScore

  Name  As String * 10
  Score As Long
  
End Type

Type typ_GameConfig

  FPS         As Boolean
  Trails      As Boolean
  Sound       As Boolean
  Turbo       As Boolean
  
End Type

Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Const SW_SHOWNORMAL = 1
Public Const SE_ERR_NOASSOC = 31
'

Public Function MakePH(str_Type As String) As cls_PhysicsHandler
    
  Dim obj_PhysicasHandler As cls_PhysicsHandler
  Dim obj_PH As Object

  Select Case str_Type
    Case "SHIP"
    
      Set obj_PH = New cls_PHShip
      
    Case "ASTEROID"
    
      Set obj_PH = New cls_PHAsteroid
      
    Case "SHOT"
    
      Set obj_PH = New cls_PHShot
      
  End Select
  
  Set obj_PhysicasHandler = obj_PH
  Set obj_PH = Nothing
  
  Set MakePH = obj_PhysicasHandler
  
End Function

Public Function OpenDoc(str_File As String)

  Dim str_Path    As String
  Dim lng_Return  As Long
  
  lng_Return = ShellExecute(GetDesktopWindow(), "open", str_File, vbNullString, vbNullString, SW_SHOWNORMAL)
  
  If lng_Return = SE_ERR_NOASSOC Then
    
    str_Path = Space(255)
    lng_Return = GetSystemDirectory(str_Path, 255)
    str_Path = Left$(str_Path, lng_Return)

    lng_Return = ShellExecute(GetDesktopWindow(), "open", "RUNDLL32.EXE", "shell32.dll,OpenAs_RunDLL " + str_File, str_Path, SW_SHOWNORMAL)

  End If
  
End Function

Public Function GetConfig() As typ_GameConfig
  
  Dim tmp_Config    As typ_GameConfig
  Dim lng_FreeFile  As Long
  Dim str_File      As String
  
  On Error Resume Next
  
  str_File = AppPath() & "config.dat"
  
  If FileExists(str_File) Then
    
    lng_FreeFile = FreeFile()
    
    Open str_File For Random Access Read As #lng_FreeFile
      Get #lng_FreeFile, 1, tmp_Config
    Close #lng_FreeFile
  
  Else
    
    With tmp_Config
    
      If InIDE() Then .FPS = True
      
      .Sound = True
      .Trails = False
      .Turbo = False
      
    End With
    
  End If
  
  GetConfig = tmp_Config
  
End Function

Public Sub SetConfig(cfg_Config As typ_GameConfig)
  
  Dim lng_FreeFile  As Long
  Dim str_File      As String
  
  On Error Resume Next
  
  str_File = AppPath() & "config.dat"

  lng_FreeFile = FreeFile()
  
  Open str_File For Random Access Write As #lng_FreeFile
    Put #lng_FreeFile, 1, cfg_Config
  Close #lng_FreeFile

End Sub

Public Function FileExists(str_File As String) As Boolean
  
  Dim lng_FreeFile As Long
  
  On Error GoTo lbl_Abort
  
  lng_FreeFile = FreeFile()
  
  Open str_File For Input As #lng_FreeFile
  Close #lng_FreeFile
  
  FileExists = True
  
  Exit Function
    
lbl_Abort:

  FileExists = False
    
End Function

Public Function AppPath() As String
  
  AppPath = IIf(Mid$(App.Path, Len(App.Path), 1) = "\", App.Path, App.Path & "\")
  
End Function

Public Function InIDE() As Boolean

  On Error GoTo lbl_Abort

  Debug.Print (1 / 0)

  Exit Function

lbl_Abort:

  InIDE = True

End Function

Function DegreeToRadian(int_Theta As Integer) As Double
  
  DegreeToRadian = int_Theta * ((4 * Atn(1)) / 180)

End Function

Function RadianToDegree(dbl_X As Double) As Integer
  
  RadianToDegree = dbl_X * (180 / (4 * Atn(1)))

End Function

Function GetDistance(tmp_PointA As typ_Point, tmp_PointB As typ_Point) As Integer

'   This function uses the pathagorean theorem to find the
' distance between the centers of two different rectangles.
' It returns an integer value for the distance.


  Dim int_DeltaX As Long            ' X difference of
                                    ' rectangles.
                                    
  Dim int_DeltaY As Long            ' Y difference of
                                    ' rectangles.
    
  Dim lng_Distance As Long          ' Distance between
                                    ' centers of
                                    ' rectangles.
  
  int_DeltaX = Abs(tmp_PointA.X - tmp_PointB.X)
  int_DeltaY = Abs(tmp_PointA.Y - tmp_PointB.Y)
                                    ' Find the X and Y
                                    ' differences between
                                    ' the two points.
                                    
  lng_Distance = Sqr(int_DeltaX * int_DeltaX + int_DeltaY * int_DeltaY)
                                    ' Pathagorean theorem,
                                    ' a² + b² = c²
                                    
  GetDistance = CInt(lng_Distance)  ' Returns distance.

End Function

Function GetAngle(PointA As typ_Point, PointB As typ_Point) As Integer
  
'   This function takes two rectangles and finds the angle
' between them using standard angle position.  For example,
' right is 0° or 360° , up is 90°, left is 180°, and down
' is 270°.  It uses the arctangent of the ratio of the X
' and Y values.  The value is converted to degree measure
' from radian.  The points are analyzed in order to
' determine which quadrants they are in and returns a
' value based on that.

  
  Dim int_DeltaX As Long            ' X difference of
                                    ' rectangles.
                                    
  Dim int_DeltaY As Long            ' Y difference of
                                    ' rectangles.
                                    
  Dim tmp_PointA As typ_Point
  Dim tmp_PointB As typ_Point
  
  
  With tmp_PointA
    .X = Int(PointA.X)
    .Y = Int(PointA.Y)
  End With
  
  With tmp_PointB
    .X = Int(PointB.X)
    .Y = Int(PointB.Y)
  End With
  
'-Top Left Quadrant------------------------------------------
  
  If tmp_PointA.X < tmp_PointB.X And tmp_PointA.Y < tmp_PointB.Y Then
    
    int_DeltaX = tmp_PointB.X - tmp_PointA.X
    int_DeltaY = tmp_PointB.Y - tmp_PointA.Y

    GetAngle = 360 - RadianToDegree(Atn(int_DeltaY / int_DeltaX))
     
  End If
'------------------------------------------------------------

'-Top Right Quadrant-----------------------------------------
  
  If tmp_PointA.X > tmp_PointB.X And tmp_PointA.Y < tmp_PointB.Y Then
      
    int_DeltaX = tmp_PointA.X - tmp_PointB.X
    int_DeltaY = tmp_PointB.Y - tmp_PointA.Y

    GetAngle = 180 + RadianToDegree(Atn(int_DeltaY / int_DeltaX))
    
  End If
  
'------------------------------------------------------------

'-Bottom Right Quadrant--------------------------------------
  
  If tmp_PointA.X > tmp_PointB.X And tmp_PointA.Y > tmp_PointB.Y Then
    
    int_DeltaX = tmp_PointA.X - tmp_PointB.X
    int_DeltaY = tmp_PointA.Y - tmp_PointB.Y

    GetAngle = 180 - RadianToDegree(Atn(int_DeltaY / int_DeltaX))
    
  End If

'------------------------------------------------------------
  
'-Bottom Left Quadrant---------------------------------------
  
  If tmp_PointA.X < tmp_PointB.X And tmp_PointA.Y > tmp_PointB.Y Then
    
    int_DeltaX = tmp_PointB.X - tmp_PointA.X
    int_DeltaY = tmp_PointA.Y - tmp_PointB.Y

    GetAngle = RadianToDegree(Atn(int_DeltaY / int_DeltaX))
    
  End If
  
'------------------------------------------------------------

End Function

Public Function GetRect(int_Left As Integer, int_Top As Integer, int_Right As Integer, int_Bottom As Integer) As RECT
  
  Dim rct_Return As RECT
  
  With rct_Return
    .Left = int_Left
    .Top = int_Top
    .Right = int_Right
    .Bottom = int_Bottom
  End With
  
  GetRect = rct_Return
  
End Function

Public Function Collide(RectA As RECT, RectB As RECT) As Boolean
    
    Dim rct_Return As RECT
    
    Collide = IntersectRect(rct_Return, RectA, RectB)

End Function

