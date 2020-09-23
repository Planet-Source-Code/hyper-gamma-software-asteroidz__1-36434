VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_Bitmaps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private obj_BmpCollection As Collection
'

Public Property Get Item(str_Name As String) As cls_Bitmap

  On Error GoTo lbl_NotFound
  
  Set Item = obj_BmpCollection(UCase(str_Name))
  
  Exit Property
    
lbl_NotFound:

  Set Item = Nothing
    
End Property

Public Property Get Count() As Integer
  
  Count = obj_BmpCollection.Count
  
End Property

Public Function Add(str_Name As String) As cls_Bitmap

  Dim tmp_Item As cls_Bitmap
  
  On Error Resume Next
  Set tmp_Item = obj_BmpCollection.Item(UCase(str_Name))
  On Error GoTo 0
  
  If (tmp_Item Is Nothing) Then
  
    Set tmp_Item = New cls_Bitmap
    Set tmp_Item.Parent = Me
        
    obj_BmpCollection.Add tmp_Item, UCase(str_Name)
    Set Add = tmp_Item
    Set tmp_Item = Nothing
  
  Else
  
    Set Add = obj_BmpCollection.Item(UCase(str_Name))
  
  End If
  
End Function

Private Sub Class_Initialize()
  
  Set obj_BmpCollection = New Collection
    
End Sub

Private Sub Class_Terminate()

  With obj_BmpCollection
    Do While .Count > 0
      .Remove .Count
    Loop
  End With
  
  Set obj_BmpCollection = Nothing
  
End Sub
