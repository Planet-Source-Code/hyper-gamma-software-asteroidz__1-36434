VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_Res"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
  
  Option Explicit
  
  Private Const WM_DISPLAYCHANGE = &H7E
  Private Const HWND_BROADCAST = &HFFFF&
  Private Const EWX_LOGOFF = 0
  Private Const EWX_SHUTDOWN = 1
  Private Const EWX_REBOOT = 2
  Private Const EWX_FORCE = 4
  Private Const CCDEVICENAME = 32
  Private Const CCFORMNAME = 32
  Private Const DM_BITSPERPEL = &H40000
  Private Const DM_PELSWIDTH = &H80000
  Private Const DM_PELSHEIGHT = &H100000
  Private Const CDS_UPDATEREGISTRY = &H1
  Private Const CDS_TEST = &H4
  Private Const DISP_CHANGE_SUCCESSFUL = 0
  Private Const DISP_CHANGE_RESTART = 1

  Private Type DEVMODE
    dmDeviceName As String * CCDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
  End Type
  
  
  Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
  Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long
  Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
  Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
  
  
  Private lng_OldWidth  As Long
  Private lng_OldHeight As Long
  Private lng_OldBPP    As Long
  Private bln_Changed   As Boolean
  '
  
Public Function ChangeRes(ByVal lng_Width As Long, ByVal lng_Height As Long, Optional ByVal lng_BPP As Long) As Boolean
  
  Dim DevM        As DEVMODE
  Dim msg_Result  As VbMsgBoxResult
  
  Dim lng_ScInfo  As Long
  Dim lng_Return  As Long
  
  On Error GoTo lbl_Abort
  
  ChangeRes = True
  
  lng_Return = EnumDisplaySettings(0&, 0&, DevM)
  If (lng_Return = 0) Then GoTo lbl_Abort
  
  With DevM
    .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
    .dmPelsWidth = lng_Width
    .dmPelsHeight = lng_Height
    .dmBitsPerPel = IIf(lng_BPP > 0, lng_BPP, lng_OldBPP)
  End With
  
  lng_Return = ChangeDisplaySettings(DevM, CDS_TEST)
  
  Select Case lng_Return
    
    Case DISP_CHANGE_RESTART
    
      msg_Result = MsgBox("In order for these changes to take place, your machine needs to restart." & vbNewLine & vbNewLine & "Would you like to restart?", vbYesNo + vbSystemModal, "Info")
      If msg_Result = vbYes Then lng_Return = ExitWindowsEx(EWX_REBOOT, 0&)
    
    Case DISP_CHANGE_SUCCESSFUL
    
      lng_Return = ChangeDisplaySettings(DevM, CDS_UPDATEREGISTRY)
      
      lng_ScInfo = lng_Width * 2 ^ 16 + lng_Height
      
      lng_Return = SendMessage(HWND_BROADCAST, WM_DISPLAYCHANGE, ByVal lng_BPP, ByVal lng_ScInfo)
    
    Case Else
    
      GoTo lbl_Abort
      
  End Select
  
  bln_Changed = True

  Exit Function
  
lbl_Abort:
  
  ChangeRes = False
  
End Function

Public Sub RestoreRes()

  Dim bln_Success As Boolean

  bln_Success = ChangeRes(lng_OldWidth, lng_OldHeight, lng_OldBPP)
  If (Not bln_Success) Then MsgBox "There was an error while attempting to restore your display settings.", vbOKOnly + vbCritical, "cls_Res::Class_Terminate"

End Sub

Private Sub Class_Initialize()
    
  Dim DevM          As DEVMODE

  Dim lng_ScreenDC  As Long
  Dim lng_Return    As Long
  
  
  On Error GoTo lbl_Abort
  
  lng_Return = EnumDisplaySettings(0&, -1, DevM)
  If (lng_Return = 0) Then GoTo lbl_Abort
  
  With DevM
    lng_OldWidth = .dmPelsWidth
    lng_OldHeight = .dmPelsHeight
    lng_OldBPP = .dmBitsPerPel
  End With
    
  Exit Sub
  
lbl_Abort:
  
  MsgBox "There was an error while attempting to access your display settings.", vbOKOnly + vbCritical, "cls_Res::Class_Initialize"
    
End Sub

Private Sub Class_Terminate()
    
  If bln_Changed Then RestoreRes

End Sub
