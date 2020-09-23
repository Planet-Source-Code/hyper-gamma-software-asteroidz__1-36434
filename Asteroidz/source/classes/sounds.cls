VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_Sounds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private obj_WavCollection As Collection
'

Public Property Get Item(str_Name As String) As cls_Sound

  On Error GoTo lbl_NotFound
  
  Set Item = obj_WavCollection(UCase(str_Name))
  
  Exit Property
    
lbl_NotFound:

  Set Item = Nothing
    
End Property

Public Function Add(str_Name As String) As cls_Sound

  Dim tmp_Item As cls_Sound
  
  On Error Resume Next
  Set tmp_Item = obj_WavCollection.Item(UCase(str_Name))
  On Error GoTo 0
  
  If (tmp_Item Is Nothing) Then
  
    Set tmp_Item = New cls_Sound
    Set tmp_Item.Parent = Me
        
    obj_WavCollection.Add tmp_Item, UCase(str_Name)
    
    Set Add = tmp_Item
    Set tmp_Item = Nothing
  
  Else
  
    Set Add = obj_WavCollection.Item(UCase(str_Name))
  
  End If
  
End Function

Private Sub Class_Initialize()
  
  Set obj_WavCollection = New Collection
    
End Sub

Private Sub Class_Terminate()

  With obj_WavCollection
    Do While .Count > 0
      .Remove .Count
    Loop
  End With
  
  Set obj_WavCollection = Nothing
  
End Sub
