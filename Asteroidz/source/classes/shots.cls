VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_Shots"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"cFrameTemplate"
Attribute VB_Ext_KEY = "Member0" ,"cFrameTemplate"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Parent As cls_Main
Private shot_Collection As Collection
'

Public Function Hit(obj_Asteroid As cls_Asteroid) As cls_Shot
    
  Dim tmp_Item    As cls_Shot
  
  Dim rct_Return  As RECT
  Dim rct_RectA   As RECT
  Dim rct_RectB   As RECT
    
  With obj_Asteroid
    rct_RectA = GetRect(.Left, .Top, .Left + .Width, .Top + .Height)
  End With
  
  For Each tmp_Item In shot_Collection
  
    With tmp_Item
                
      rct_RectB = GetRect(.Left, .Top, .Left + .Width, .Top + .Height)
      
      If IntersectRect(rct_Return, rct_RectA, rct_RectB) Then
        If Parent.Sound Then Parent.Sounds.Item("HIT").PlaySound
        Set Hit = tmp_Item
        Exit For
      End If
           
    End With
    
  Next
    
End Function

Public Function Add(str_Type As String) As cls_Shot
  
  Dim tmp_Item As cls_Shot
  
  Set tmp_Item = New cls_Shot
  
  With tmp_Item
  
    Set .Parent = Me
    
    Select Case str_Type
    
      Case "IONPULSE"
      
        Set .Bitmap = Parent.Bitmaps.Item("IONPULSE")
        .ShotType = str_Type
        .Power = 50
        .Speed = 15
        
      Case "SEEKER"
      
        Set .Bitmap = Parent.Bitmaps.Item("SEEKER")
        .ShotType = str_Type
        .Power = 150
        .Speed = 5
        
    End Select
    
    .Activate
    
  End With
  
  If Parent.Sound And (Parent.Asteroids.Count > 0) Then Parent.Sounds.Item("SHOOT").PlaySound

  shot_Collection.Add tmp_Item
  
  Set Add = tmp_Item
  Set tmp_Item = Nothing
  
End Function

Public Property Get Item(nIndex As Long) As cls_Shot
Attribute Item.VB_UserMemId = 0

  Set Item = shot_Collection(nIndex)
    
End Property

Public Property Get Count() As Long

  Count = shot_Collection.Count
    
End Property

Public Sub Remove(nIndex As Long)

  shot_Collection.Remove nIndex
    
End Sub

Public Sub Update()

  Dim tmp_Item As cls_Shot
  Dim int_Index As Integer
    
  If (Parent.Asteroids.Count = 0) Then Clear

  For Each tmp_Item In shot_Collection
    tmp_Item.Update
  Next
  
  int_Index = 1
  
  Do Until int_Index > shot_Collection.Count
    If shot_Collection.Item(int_Index).Terminating Then
      shot_Collection.Remove int_Index
    Else
      int_Index = int_Index + 1
    End If
  Loop

End Sub

Public Sub Render()

  Dim tmp_Item As cls_Shot
  
  For Each tmp_Item In shot_Collection
    tmp_Item.Render
  Next

End Sub


Public Sub Clear()

  Do While shot_Collection.Count > 0
    shot_Collection.Remove shot_Collection.Count
  Loop
  
  Set shot_Collection = New Collection
    
End Sub

Private Sub Class_Initialize()

  Set shot_Collection = New Collection
    
End Sub

Private Sub Class_Terminate()

  Do While shot_Collection.Count > 0
    shot_Collection.Remove shot_Collection.Count
  Loop
  
  Set shot_Collection = Nothing
    
End Sub

