VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_Labels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private obj_Parent        As cls_Main
Private obj_Bitmap        As cls_Bitmap

Private obj_lblCollection As Collection

Private lng_DestDC        As Long
'

Property Set Parent(obj_NewValue As cls_Main)
  
  Set obj_Parent = obj_NewValue
  
End Property

Property Get Parent() As cls_Main
  
  Set Parent = obj_Parent
  
End Property

Property Get Bitmap() As cls_Bitmap
  
  Set Bitmap = obj_Bitmap
  
End Property

Property Set Bitmap(obj_NewValue As cls_Bitmap)

  Set obj_Bitmap = obj_NewValue
  
End Property

Property Let DestDC(lng_NewValue As Long)
  
  lng_DestDC = lng_NewValue
  
End Property

Property Get DestDC() As Long
  
  DestDC = lng_DestDC
  
End Property

Public Property Get Item(str_Name As String) As cls_Label

  On Error GoTo lbl_NotFound
  
  Set Item = obj_lblCollection(UCase(str_Name))
  
  Exit Property
    
lbl_NotFound:

  Set Item = Nothing
    
End Property

Public Function Add(str_Name As String) As cls_Label

  Dim tmp_Item As cls_Label
  
  On Error Resume Next
  Set tmp_Item = obj_lblCollection.Item(UCase(str_Name))
  On Error GoTo 0
  
  If (tmp_Item Is Nothing) Then
  
    Set tmp_Item = New cls_Label
    Set tmp_Item.Parent = Me
        
    obj_lblCollection.Add tmp_Item, UCase(str_Name)
    
    Set Add = tmp_Item
    Set tmp_Item = Nothing
  
  Else
  
    Set Add = obj_lblCollection.Item(UCase(str_Name))
  
  End If
  
End Function

Public Sub Render()

  Dim tmp_Item As cls_Label
  
  For Each tmp_Item In obj_lblCollection
    tmp_Item.Render
  Next

End Sub

Private Sub Class_Initialize()
  
  Set obj_lblCollection = New Collection
    
End Sub

Private Sub Class_Terminate()

  With obj_lblCollection
    Do While .Count > 0
      .Remove .Count
    Loop
  End With
  
  Set obj_lblCollection = Nothing
  
End Sub
