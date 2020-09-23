VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_Effects"
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
  Private fx_Collection As Collection
  '
  
Public Function Add(str_Type As String) As cls_Effect

  Dim tmp_Item As cls_Effect
  
  Set tmp_Item = New cls_Effect
  
  With tmp_Item
  
    Set .Parent = Me
    
    fx_Collection.Add tmp_Item
    
    .EffectType = str_Type
    
    Set .Bitmap = Parent.Bitmaps.Item(str_Type)
        
  End With
  
  Set Add = tmp_Item

  Set tmp_Item = Nothing
    
End Function

Public Property Get Item(int_Index As Integer) As cls_Effect

  Set Item = fx_Collection(int_Index)
    
End Property

Public Property Get Count() As Integer

  Count = fx_Collection.Count
    
End Property

Public Sub Remove(int_Index As Integer)

  fx_Collection.Remove int_Index
    
End Sub

Public Sub Update()

  Dim tmp_Item As cls_Effect
  
  Dim int_IndexCounter As Integer

  int_IndexCounter = 1
  
  For Each tmp_Item In fx_Collection
    tmp_Item.Update
    If fx_Collection.Item(int_IndexCounter).Terminating Then
      fx_Collection.Remove int_IndexCounter
    Else
      int_IndexCounter = int_IndexCounter + 1
    End If
  Next
    
End Sub

Public Sub Render()

  Dim tmp_Item As cls_Effect
  
  For Each tmp_Item In fx_Collection
    tmp_Item.Render
  Next

End Sub

Public Sub Clear()
  
  With fx_Collection
    Do While .Count > 0
      .Remove .Count
    Loop
  End With
  
  Set fx_Collection = New Collection
    
End Sub

Private Sub Class_Initialize()

  Set fx_Collection = New Collection
  
End Sub

Private Sub Class_Terminate()

  With fx_Collection
    Do While .Count > 0
      .Remove .Count
    Loop
  End With

  Set fx_Collection = Nothing
    
End Sub


