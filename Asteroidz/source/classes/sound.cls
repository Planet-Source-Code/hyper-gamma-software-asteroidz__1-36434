VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_Sound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (lpszSoundName As Byte, ByVal uFlags As Long) As Long

Private Const SND_ASYNC = &H1
Private Const SND_NODEFAULT = &H2
Private Const SND_MEMORY = &H4
Private Const SND_LOOP = &H8

Private obj_Parent As cls_Sounds
Private byt_Sound() As Byte
'

Public Property Get Parent() As cls_Sounds
  
  Set Parent = obj_Parent
  
End Property

Public Property Set Parent(obj_NewValue As cls_Sounds)
  
  Set obj_Parent = obj_NewValue
  
End Property

Public Function LoadSoundFromRes(SoundID As Variant, Optional str_Type As String = "CUSTOM") As Boolean
  
  Dim byt_Array() As Byte
  Dim lng_IndexCounter As Long
  Dim lng_FileLen As Long
  
  On Error GoTo lbl_Abort
  
  LoadSoundFromRes = True
  
  byt_Array = LoadResData(SoundID, str_Type)
  
  lng_FileLen = UBound(byt_Array)
  ReDim byt_Sound(lng_FileLen)
  
  For lng_IndexCounter = 0 To lng_FileLen
    byt_Sound(lng_IndexCounter) = byt_Array(lng_IndexCounter)
  Next
  
  Exit Function
  
lbl_Abort:
  
  LoadSoundFromRes = False
  
End Function

Public Function LoadSoundFromFile(str_FileName As String) As Boolean

  Dim int_FreeFile As Integer
  Dim str_File As String
  Dim lng_IndexCounter As Long
  Dim lng_FileLen As Long
  
  On Error GoTo lbl_Abort

  int_FreeFile = FreeFile
  
  Open str_FileName For Binary Access Read As #int_FreeFile
  
    lng_FileLen = LOF(int_FreeFile)
    str_File = Space(lng_FileLen)
    ReDim byt_Sound(lng_FileLen)
    
    Get #int_FreeFile, , str_File
    
  Close #int_FreeFile
  
  For lng_IndexCounter = 0 To lng_FileLen - 1
    byt_Sound(lng_IndexCounter) = Asc(Mid(str_File, lng_IndexCounter + 1))
  Next
  
  Exit Function
  
lbl_Abort:
  
  LoadSoundFromFile = False
  
End Function

Public Sub PlaySound()
  
  On Error Resume Next
  
  sndPlaySound byt_Sound(0), SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY
  
End Sub

Public Sub LoopSound()
  
  On Error Resume Next
  
  sndPlaySound byt_Sound(0), SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY Or SND_LOOP

End Sub

Public Sub StopSound()
  
  On Error Resume Next
  
  sndPlaySound 0, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY

End Sub


