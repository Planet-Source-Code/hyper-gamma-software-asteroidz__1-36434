VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_Buttons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Parent           As frm_Splash

Public Bitmap           As cls_Bitmap
Public CurrentButton    As cls_Button

Private mnu_Collection  As Collection
'

Public Sub Activate()
  
  With Add("PLAYGAME")
    .Left = 175
    .Top = 170
    .Frame = 0
    .Activate
  End With
  
  With Add("VIEWSCORES")
    .Left = 175
    .Top = 195
    .Frame = 1
    .Activate
  End With
  
  With Add("VIEWREADME")
    .Left = 175
    .Top = 220
    .Frame = 2
    .Activate
  End With
  
  With Add("QUIT")
    .Left = 175
    .Top = 245
    .Frame = 3
    .Activate
  End With
    
End Sub

Public Sub Render()

  Dim tmp_Item As cls_Button
  
  For Each tmp_Item In mnu_Collection
    tmp_Item.Render
  Next
  
End Sub

Public Sub Update()
  
  Dim tmp_Item As cls_Button
  Dim rct_Return As RECT
  Dim bln_MenuSelected As Boolean
   
  Static lng_StartTime As Long
           
  For Each tmp_Item In mnu_Collection
        
    If tmp_Item.MouseOver(Parent.MouseX, Parent.MouseY) Then
      If ObjPtr(CurrentButton) Then CurrentButton.Enabled = False
      Set CurrentButton = tmp_Item
      CurrentButton.Enabled = True
      bln_MenuSelected = True
    End If
    
  Next
    
  If bln_MenuSelected Then
    If Parent.MousePressed Then Parent.Command CurrentButton.Command
  Else
    If ObjPtr(CurrentButton) Then CurrentButton.Enabled = False
  End If
      
End Sub

Private Function Add(str_Name As String) As cls_Button
  
  Dim tmp_Item As cls_Button
  
  On Error Resume Next
  Set tmp_Item = mnu_Collection.Item(str_Name)
  On Error GoTo 0
  
  If (tmp_Item Is Nothing) Then

    Set tmp_Item = New cls_Button
    Set tmp_Item.Parent = Me
    tmp_Item.Command = str_Name
    
    mnu_Collection.Add tmp_Item, str_Name
    
    Set Add = tmp_Item
    Set tmp_Item = Nothing
  
  Else
  
    Set Add = mnu_Collection.Item(str_Name)
  
  End If
  
End Function

Private Sub Class_Initialize()

  Set mnu_Collection = New Collection
    
End Sub

Private Sub Class_Terminate()

  With mnu_Collection
    Do While .Count > 0
      .Remove .Count
    Loop
  End With
  
  Set mnu_Collection = Nothing
    
End Sub
