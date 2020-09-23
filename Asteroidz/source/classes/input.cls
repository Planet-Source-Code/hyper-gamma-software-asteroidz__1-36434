VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_Input"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'Project Info...
' - Author:          Zach "Orion" Collette
' - Company:         Hyper-Gamma Software
' - Contact:         kingtheoden17@hotmail.com
' - Description:     Keyboard Input Class
' - Test Platform:   Windows 98 SE
' - Processor:       P3 450MHz.
'------------------------------------------------------------
 
 
'   This class is designed to emulate key events similar
' to those used in forms.  It should be used in either a
' timer or loop (be sure to us 'DoEvents()').  If it is
' not "refreshed" by the 'Refresh()' event, then you won't
' get an accurate result.  Everything is totally
' encapsulated, so all you do is add it to your project,
' and create an input object where it needs to be.

'   Feel free to use this and distribute it as well, but
' as with all submissions, please do not alter in any
' way or take credit for this work.  Thank You... Enjoy


  ' Enforce variable declaration.
Option Explicit

  ' Structure used to hold key data.
Private Type typ_Key
  
  Key As KeyCodeConstants  ' KeyCode used by Windows
  
  OldValue As Boolean      ' Was it pressed last time
                           ' key was Refreshd?
  
  CurrentValue As Boolean  ' Is it pressed now?
  
End Type
  
  
  ' API function, returns true if the key is pressed.
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
  

  ' Our array of key_Array.
Private key_Array() As typ_Key

  ' Our key count.
Private int_KeyCount As Integer
'

Public Sub AddKey(ByVal Key As KeyCodeConstants)
        
    ' Adds a key to the array.
  ReDim Preserve key_Array(int_KeyCount)

    ' Sets the KeyCode to the key specified
  key_Array(int_KeyCount).Key = Key
    
    ' Increase the count.
  int_KeyCount = int_KeyCount + 1
    
End Sub

Public Sub Refresh()
  
  Dim int_KeyCounter As Integer  ' Used to loop through
                                 ' key array.
  
    'Cycle through all the key_Array.
  For int_KeyCounter = 0 To (int_KeyCount - 1)
    
    With key_Array(int_KeyCounter)
    
        ' Sets the old value to what the current value
        ' is before refreshing current value.
      .OldValue = .CurrentValue
      
        ' Refresh the current value.
      .CurrentValue = GetAsyncKeyState(.Key)
      
    End With
    
  Next
    
End Sub

Public Function KeyDown(ByVal Key As KeyCodeConstants) As Boolean
  
  Dim int_KeyCounter As Integer  ' Used to loop through
                                 ' key array.
  
    'Cycle through all the key_Array.
  For int_KeyCounter = 0 To (int_KeyCount - 1)
    
    With key_Array(int_KeyCounter)
    
        ' If key wasn't pressed before it was refreshed,
        ' and is now, then the key was just pressed.
      If .Key = Key And ((Not .OldValue) And .CurrentValue) Then KeyDown = True
    
    End With
  
  Next
  
End Function

Public Function KeyPressed(ByVal Key As KeyCodeConstants) As Boolean
  
  Dim int_KeyCounter As Integer  ' Used to loop through
                                 ' key array.
  
    'Cycle through all the key_Array.
  For int_KeyCounter = 0 To (int_KeyCount - 1)
    
    With key_Array(int_KeyCounter)
      
        ' If key was pressed before it was refreshed,
        ' and is now, then the key is pressed.
      If .Key = Key And (.OldValue And .CurrentValue) Then KeyPressed = True
    
    End With
  
  Next

End Function

Public Function KeyUp(ByVal Key As KeyCodeConstants) As Boolean
  
  Dim int_KeyCounter As Integer  ' Used to loop through
                                 ' key array.
  
    'Cycle through all the key_Array.
  For int_KeyCounter = 0 To (int_KeyCount - 1)
    
    With key_Array(int_KeyCounter)
    
        ' If key was pressed before it was refreshed,
        ' and isn't now, then the key was just released.
      If .Key = Key And (.OldValue And (Not .CurrentValue)) Then KeyUp = True
    
    End With
    
  Next

End Function
