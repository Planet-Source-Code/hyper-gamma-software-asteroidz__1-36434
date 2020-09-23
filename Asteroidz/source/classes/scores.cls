VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_Scores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private Type typ_HighScore

  Name  As String * 10
  Score As Long
  
End Type

Public NewName          As String

Public Sub CheckHighScore(lng_Score As Long)
  
  Dim Scores(1 To 11)   As typ_HighScore
  Dim HighScore         As typ_HighScore
  
  Dim int_ScoreCounter  As Integer
  Dim str_HighScore     As String
  Dim str_Temp()        As String
  Dim lng_FreeFile      As Long
  Dim bln_NewScore      As Boolean
  Dim int_NewScoreIndex As Integer
  Dim str_File          As String
  
  On Error Resume Next
  
  str_File = AppPath() & "scores.dat"
  
  If FileExists(str_File) Then
    
    lng_FreeFile = FreeFile()

    Open str_File For Random Access Read As #lng_FreeFile Len = 14
      For int_ScoreCounter = 1 To 10
        Get #lng_FreeFile, int_ScoreCounter, Scores(int_ScoreCounter)
      Next
    Close #lng_FreeFile
  
  Else
  
    For int_ScoreCounter = 1 To 10
      
      str_HighScore = LoadResString(int_ScoreCounter - 1)
      str_Temp = Split(str_HighScore, "|")
      
      Scores(int_ScoreCounter).Name = str_Temp(0)
      Scores(int_ScoreCounter).Score = CLng(str_Temp(1))
    
    Next
  
  End If
  
  For int_ScoreCounter = 10 To 1 Step -1
    
    If lng_Score > Scores(int_ScoreCounter).Score Then
      int_NewScoreIndex = int_ScoreCounter
      bln_NewScore = True
    End If
    
  Next
  
  If int_NewScoreIndex <> 0 Then
    
    For int_ScoreCounter = 10 To int_NewScoreIndex Step -1
      Scores(int_ScoreCounter + 1) = Scores(int_ScoreCounter)
    Next
    
    Load frm_NewHighScore
    
    frm_NewHighScore.Show vbModal
    
    If NewName = "" Then Exit Sub
    
    Scores(int_NewScoreIndex).Name = NewName
    Scores(int_NewScoreIndex).Score = lng_Score
  
  End If
  
  lng_FreeFile = FreeFile()
  
  Open str_File For Random Access Write As #lng_FreeFile Len = 14
    
    For int_ScoreCounter = 1 To 10
      HighScore = Scores(int_ScoreCounter)
      Put #lng_FreeFile, int_ScoreCounter, HighScore
    Next
    
  Close #lng_FreeFile
  
  If bln_NewScore Then ShowHighScores
    
End Sub

Public Sub ShowHighScores()
  
  Dim HighScore         As typ_HighScore
  Dim Scores(1 To 11)   As typ_HighScore
  
  Dim NameList          As String
  Dim ScoreList         As String
  Dim str_Message       As String
  Dim bln_FileExists    As Boolean
  Dim int_ScoreCounter  As Integer
  Dim str_HighScore     As String
  Dim str_Temp()        As String
  Dim lng_FreeFile      As Long
  Dim str_File          As String
  
  str_File = AppPath() & "scores.dat"
  
  bln_FileExists = FileExists(str_File)
  
  If Not bln_FileExists Then
  
    For int_ScoreCounter = 1 To 10
      
      str_HighScore = LoadResString(int_ScoreCounter - 1)
      str_Temp = Split(str_HighScore, "|")
      
      Scores(int_ScoreCounter).Name = str_Temp(0)
      Scores(int_ScoreCounter).Score = CLng(str_Temp(1))
    
    Next
    
    lng_FreeFile = FreeFile()
    
    Open str_File For Random Access Write As #lng_FreeFile Len = 14
            
      For int_ScoreCounter = 1 To 10
        HighScore = Scores(int_ScoreCounter)
        Put #lng_FreeFile, int_ScoreCounter, HighScore
      Next
      
    Close #lng_FreeFile
    
  End If
  
  lng_FreeFile = FreeFile()
  
  Open str_File For Random Access Read As #lng_FreeFile Len = 14
    
    For int_ScoreCounter = 1 To 10
      Get #lng_FreeFile, int_ScoreCounter, HighScore
      NameList = NameList & Format((int_ScoreCounter), "00") & ") " & HighScore.Name & vbNewLine
      ScoreList = ScoreList & Format(HighScore.Score, "000000") & vbNewLine
    Next
        
  Close #lng_FreeFile
  
  Load frm_HighScores
  
  With frm_HighScores
    .NameList = NameList
    .ScoreList = ScoreList
    .Render
    .Show vbModal
  End With
  
End Sub

