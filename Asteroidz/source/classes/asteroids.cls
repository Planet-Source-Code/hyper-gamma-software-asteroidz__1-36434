VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_Asteroids"
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
  
  Public Parent As cls_Main
  
  Private nme_Collection As Collection
  '
  
Public Property Get Count() As Integer
  
  Count = nme_Collection.Count
  
End Property

Private Sub Shatter(obj_Parent As cls_Asteroid)
  
  Dim int_AsteroidCounter As Integer
  Dim int_AsteroidType As Integer
  Dim str_AsteroidType As String
  
  int_AsteroidType = CInt(Right(obj_Parent.EnemyType, 1))
  If int_AsteroidType = 4 Then Exit Sub

  str_AsteroidType = "ASTEROID" & CStr(int_AsteroidType + 1)
  
  For int_AsteroidCounter = 0 To 1
    With Add(str_AsteroidType)
      .Angle = Rnd * 360
      .Speed = (Rnd * 3) + 1
      .Activate
      .Left = obj_Parent.Left + ((obj_Parent.Width - .Width) / 2)
      .Top = obj_Parent.Top + ((obj_Parent.Height - .Height) / 2)
    End With
  Next

End Sub

Public Function WithinRange(int_X As Integer, int_Y As Integer, int_Range As Integer) As cls_Asteroid
  
  Dim tmp_Item As cls_Asteroid
  
  Dim PointA As typ_Point
  Dim PointB As typ_Point
  
  With PointA
    .X = int_X
    .Y = int_Y
  End With
    
  For Each tmp_Item In nme_Collection
  
    With PointB
      .X = tmp_Item.Left + (tmp_Item.Width / 2)
      .Y = tmp_Item.Top + (tmp_Item.Height / 2)
    End With
    
    If GetDistance(PointA, PointB) < int_Range Then Set WithinRange = tmp_Item
  
  Next

End Function

Public Function Hit(obj_Asteroid As cls_Asteroid) As cls_Asteroid
    
  Dim tmp_Item    As cls_Asteroid
  
  Dim rct_Return  As RECT
  Dim rct_RectA   As RECT
  Dim rct_RectB   As RECT
    
  With obj_Asteroid
    
    rct_RectA = GetRect(.Left, .Top, .Left + .Width, .Top + .Height)
  
  End With
  
  For Each tmp_Item In nme_Collection
  
    With tmp_Item
    
      rct_RectB = GetRect(.Left, .Top, .Left + .Width, .Top + .Height)
      
      If IntersectRect(rct_Return, rct_RectA, rct_RectB) Then
        If .Index <> obj_Asteroid.Index Then
          Set Hit = tmp_Item
          Exit For
        End If
      End If
            
    End With
    
  Next
    
End Function

Public Sub Activate(int_Level As Integer)
  
  Dim int_AsteroidCounter As Integer
    
  For int_AsteroidCounter = 0 To (int_Level \ 2)
    With Add("ASTEROID1")
      .Left = Rnd * Parent.Parent.ScaleWidth
      .Top = Rnd * Parent.Parent.ScaleHeight
      .Angle = Rnd * 360
      .Speed = (Rnd * 3) + 1
      .Activate
    End With
  Next
  
End Sub

Public Function Add(str_Type As String) As cls_Asteroid

  On Error GoTo Err_AlreadyDefined
  
  Dim tmp_Item As cls_Asteroid
  Static lng_IndexCounter As Long
  
  Set tmp_Item = New cls_Asteroid
  
  With tmp_Item
   
  Set .Parent = Me
   
  .EnemyType = str_Type

    Select Case str_Type
      Case "ASTEROID1"
        Set .Bitmap = Parent.Bitmaps.Item("ASTEROID1")
        .Index = lng_IndexCounter
        .Shield = 400
        .Power = 20
      Case "ASTEROID2"
        Set .Bitmap = Parent.Bitmaps.Item("ASTEROID2")
        .Index = lng_IndexCounter
        .Shield = 300
        .Power = 15
      Case "ASTEROID3"
        Set .Bitmap = Parent.Bitmaps.Item("ASTEROID3")
        .Index = lng_IndexCounter
        .Shield = 200
        .Power = 10
      Case "ASTEROID4"
        Set .Bitmap = Parent.Bitmaps.Item("ASTEROID4")
        .Index = lng_IndexCounter
        .Shield = 100
        .Power = 5
     End Select
   
   End With
  
  nme_Collection.Add tmp_Item
  
  lng_IndexCounter = lng_IndexCounter + 1
  
  Set Add = tmp_Item
  Set tmp_Item = Nothing

Err_AlreadyDefined:

End Function

Public Sub Update()

  Dim tmp_Item          As cls_Asteroid
  Dim int_IndexCounter  As Integer
  Dim str_Effect        As String
  Dim int_Effect        As Integer
  Dim int_Asteroid      As Integer
  
  int_IndexCounter = 1
  
  For Each tmp_Item In nme_Collection
  
    tmp_Item.Update
    
    If tmp_Item.Dying Then
    
      If Parent.Sound Then Parent.Sounds.Item("EXPLOSION").PlaySound
      
      int_Asteroid = CInt(Right(tmp_Item.EnemyType, 1))
      int_Effect = Round((int_Asteroid / 2) + 0.1)
      str_Effect = "EXPLOSION" & CStr(int_Effect)
      
      Parent.Score = Parent.Score + IIf(tmp_Item.ShotByPlayer, ((5 - int_Asteroid) * 150), -(((5 - int_Asteroid) * 150) / 2))
    
      With Parent.Effects.Add(str_Effect)
        Set .ParentObject = tmp_Item
        .Left = tmp_Item.Left + (tmp_Item.Width / 2)
        .Top = tmp_Item.Top + (tmp_Item.Height / 2)
        .Activate
      End With
      
      Shatter tmp_Item
      
      tmp_Item.Terminating = True
      
    End If
    
    If tmp_Item.Terminating Then
      nme_Collection.Remove int_IndexCounter
    Else
      int_IndexCounter = int_IndexCounter + 1
    End If
    
  Next
  
End Sub

Public Sub Render()

  Dim tmp_Item As cls_Asteroid
  
  For Each tmp_Item In nme_Collection
    tmp_Item.Render
  Next

End Sub

Public Sub KillAll()

  Dim tmp_Item As cls_Asteroid
  
  For Each tmp_Item In nme_Collection
    tmp_Item.Dying = True
  Next

End Sub

Public Sub Clear()
  
  Do While nme_Collection.Count > 0
    nme_Collection.Remove nme_Collection.Count
  Loop
  
  Set nme_Collection = New Collection
    
End Sub

Private Sub Class_Initialize()

  Set nme_Collection = New Collection
    
End Sub

Private Sub Class_Terminate()

  With nme_Collection
    Do While .Count > 0
      .Remove .Count
    Loop
  End With
  
  Set nme_Collection = Nothing
    
End Sub

