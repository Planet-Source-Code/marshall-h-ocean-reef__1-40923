VERSION 5.00
Begin VB.Form frmHighScore 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "High Scores"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3690
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   164
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   246
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   1545
      TabIndex        =   16
      Top             =   2040
      Width           =   960
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2625
      TabIndex        =   15
      Top             =   2040
      Width           =   960
   End
   Begin VB.Label lblScore5 
      Alignment       =   1  'Right Justify
      Caption         =   "Score5"
      Height          =   270
      Left            =   2040
      TabIndex        =   14
      Top             =   1545
      Width           =   1455
   End
   Begin VB.Label lblScore4 
      Alignment       =   1  'Right Justify
      Caption         =   "Score4"
      Height          =   375
      Left            =   2040
      TabIndex        =   13
      Top             =   1185
      Width           =   1455
   End
   Begin VB.Label lblScore3 
      Alignment       =   1  'Right Justify
      Caption         =   "Score3"
      Height          =   285
      Left            =   2040
      TabIndex        =   12
      Top             =   825
      Width           =   1455
   End
   Begin VB.Label lblScore2 
      Alignment       =   1  'Right Justify
      Caption         =   "Score2"
      Height          =   375
      Left            =   2040
      TabIndex        =   11
      Top             =   465
      Width           =   1455
   End
   Begin VB.Label lblScore1 
      Alignment       =   1  'Right Justify
      Caption         =   "Score1"
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Top             =   105
      Width           =   1455
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   8
      X2              =   240
      Y1              =   129
      Y2              =   129
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   8
      X2              =   237
      Y1              =   128
      Y2              =   128
   End
   Begin VB.Label lblName5 
      Caption         =   "Name5"
      Height          =   375
      Left            =   405
      TabIndex        =   9
      Top             =   1560
      Width           =   1545
   End
   Begin VB.Label lblName4 
      Caption         =   "Name4"
      Height          =   375
      Left            =   405
      TabIndex        =   8
      Top             =   1200
      Width           =   1530
   End
   Begin VB.Label lblName3 
      Caption         =   "Name3"
      Height          =   255
      Left            =   405
      TabIndex        =   7
      Top             =   840
      Width           =   1545
   End
   Begin VB.Label lblName2 
      Caption         =   "Name2"
      Height          =   255
      Left            =   405
      TabIndex        =   6
      Top             =   480
      Width           =   1515
   End
   Begin VB.Label lblName1 
      Caption         =   "Name1"
      Height          =   255
      Left            =   405
      TabIndex        =   5
      Top             =   120
      Width           =   1470
   End
   Begin VB.Label Label5 
      Caption         =   "5."
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "4."
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "3."
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "2."
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "1."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "frmHighScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClear_Click()
      ClearScore ("OceanReef")
End Sub

Private Sub cmdOK_Click()
    Me.Hide
End Sub

Public Sub Form_Load()
    GetHiScore "OceanReef", True
End Sub


Function GetHiScore(AppName As String, Optional ShowScores As Boolean = False) As String

Dim h1 As Integer, h2 As Integer, h3 As Integer, h4 As Integer, h5 As Integer
Dim strH1 As String, strH2 As String, strH3 As String, strH4 As String, strH5 As String
Dim tempStr As String

h1 = GetSetting(AppName, "HiScore", "h1", "1000")
h2 = GetSetting(AppName, "HiScore", "h2", "500")
h3 = GetSetting(AppName, "HiScore", "h3", "250")
h4 = GetSetting(AppName, "HiScore", "h4", "100")
h5 = GetSetting(AppName, "HiScore", "h5", "50")
strH1 = GetSetting(AppName, "HiScore", "strH1", "Marshall")
strH2 = GetSetting(AppName, "HiScore", "strH2", "Charlotte")
strH3 = GetSetting(AppName, "HiScore", "strH3", "Yogi")
strH4 = GetSetting(AppName, "HiScore", "strH4", "Todd")
strH5 = GetSetting(AppName, "HiScore", "strH5", "Brad")

If ShowScores = True Then
        With frmHighScore
        .lblName1.Caption = strH1
        .lblName2.Caption = strH2
        .lblName3.Caption = strH3
        .lblName4.Caption = strH4
        .lblName5.Caption = strH5
        
        .lblScore1.Caption = h1
        .lblScore2.Caption = h2
        .lblScore3.Caption = h3
        .lblScore4.Caption = h4
        .lblScore5.Caption = h5
        'tempStr = "High Scores:" & vbCrLf & vbCrLf & strH1 & " - " & h1 & vbCrLf & strH2 & " - " & h2 & vbCrLf & strH3 & " - " & h3 & vbCrLf & strH4 & " - " & h4 & vbCrLf & strH5 & " - " & h5
        'MsgBox tempStr, , "Hi Scores"
        End With
End If

tempStr = "High Scores:" & vbCrLf & vbCrLf & strH1 & " - " & h1 & vbCrLf & strH2 & " - " & h2 & vbCrLf & strH3 & " - " & h3 & vbCrLf & strH4 & " - " & h4 & vbCrLf & strH5 & " - " & h5

GetHiScore = tempStr
End Function

Public Function SaveHiScore(AppName As String, Score, Optional ShowScores As Boolean = False) As String
Dim h1 As Integer, h2 As Integer, h3 As Integer, h4 As Integer, h5 As Integer
Dim strH1 As String, strH2 As String, strH3 As String, strH4 As String, strH5 As String
Dim tempStr As String

h1 = GetSetting(AppName, "HiScore", "h1", "100")
h2 = GetSetting(AppName, "HiScore", "h2", "50")
h3 = GetSetting(AppName, "HiScore", "h3", "25")
h4 = GetSetting(AppName, "HiScore", "h4", "10")
h5 = GetSetting(AppName, "HiScore", "h5", "5")
strH1 = GetSetting(AppName, "HiScore", "strH1", "Marshall")
strH2 = GetSetting(AppName, "HiScore", "strH2", "Charlotte")
strH3 = GetSetting(AppName, "HiScore", "strH3", "Yogi")
strH4 = GetSetting(AppName, "HiScore", "strH4", "Todd")
strH5 = GetSetting(AppName, "HiScore", "strH5", "Brad")

If Score > h1 Then
        NewName = InputBox("You got first place!", "New High Score!", strH1)
        SaveSetting AppName, "HiScore", "strH5", strH4
        SaveSetting AppName, "HiScore", "strH4", strH3
        SaveSetting AppName, "HiScore", "strH3", strH2
        SaveSetting AppName, "HiScore", "strH2", strH1
        SaveSetting AppName, "HiScore", "strH1", NewName
        SaveSetting AppName, "HiScore", "h5", h4
        SaveSetting AppName, "HiScore", "h4", h3
        SaveSetting AppName, "HiScore", "h3", h2
        SaveSetting AppName, "HiScore", "h2", h1
        SaveSetting AppName, "HiScore", "h1", Score
ElseIf Score > h2 Then
        NewName = InputBox("You got second place!", "New High Score!", strH2)
        SaveSetting AppName, "HiScore", "strH5", strH4
        SaveSetting AppName, "HiScore", "strH4", strH3
        SaveSetting AppName, "HiScore", "strH3", strH2
        SaveSetting AppName, "HiScore", "strH1", strH1
        SaveSetting AppName, "HiScore", "strH2", NewName
        SaveSetting AppName, "HiScore", "h5", h4
        SaveSetting AppName, "HiScore", "h4", h3
        SaveSetting AppName, "HiScore", "h3", h2
        SaveSetting AppName, "HiScore", "h2", Score
        SaveSetting AppName, "HiScore", "h1", h1
ElseIf Score > h3 Then
        NewName = InputBox("You got third place!", "New High Score!", strH3)
        SaveSetting AppName, "HiScore", "strH5", strH4
        SaveSetting AppName, "HiScore", "strH4", strH3
        SaveSetting AppName, "HiScore", "strH2", strH2
        SaveSetting AppName, "HiScore", "strH1", strH1
        SaveSetting AppName, "HiScore", "strH3", NewName
        SaveSetting AppName, "HiScore", "h5", h4
        SaveSetting AppName, "HiScore", "h4", h3
        SaveSetting AppName, "HiScore", "h3", Score
        SaveSetting AppName, "HiScore", "h2", h2
        SaveSetting AppName, "HiScore", "h1", h1
ElseIf Score > h4 Then
        NewName = InputBox("You got fourth place!", "New High Score!", strH2)
        SaveSetting AppName, "HiScore", "strH5", strH4
        SaveSetting AppName, "HiScore", "strH3", strH3
        SaveSetting AppName, "HiScore", "strH2", strH2
        SaveSetting AppName, "HiScore", "strH1", strH1
        SaveSetting AppName, "HiScore", "strH4", NewName
        SaveSetting AppName, "HiScore", "h5", h4
        SaveSetting AppName, "HiScore", "h4", Score
        SaveSetting AppName, "HiScore", "h3", h3
        SaveSetting AppName, "HiScore", "h2", h2
        SaveSetting AppName, "HiScore", "h1", h1
ElseIf Score > h5 Then
        NewName = InputBox("You got fifth place!", "New High Score!", strH2)
        SaveSetting AppName, "HiScore", "strH1", strH1
        SaveSetting AppName, "HiScore", "strH2", strH2
        SaveSetting AppName, "HiScore", "strH3", strH3
        SaveSetting AppName, "HiScore", "strH4", strH4
        SaveSetting AppName, "HiScore", "strH5", NewName
        SaveSetting AppName, "HiScore", "h5", Score
        SaveSetting AppName, "HiScore", "h4", h4
        SaveSetting AppName, "HiScore", "h3", h3
        SaveSetting AppName, "HiScore", "h2", h2
        SaveSetting AppName, "HiScore", "h1", h1
Else
        Exit Function
End If

h1 = GetSetting(AppName, "HiScore", "h1", "100")
h2 = GetSetting(AppName, "HiScore", "h2", "50")
h3 = GetSetting(AppName, "HiScore", "h3", "25")
h4 = GetSetting(AppName, "HiScore", "h4", "10")
h5 = GetSetting(AppName, "HiScore", "h5", "5")
strH1 = GetSetting(AppName, "HiScore", "strH1", "Marshall")
strH2 = GetSetting(AppName, "HiScore", "strH2", "Charlotte")
strH3 = GetSetting(AppName, "HiScore", "strH3", "Yogi")
strH4 = GetSetting(AppName, "HiScore", "strH4", "Todd")
strH5 = GetSetting(AppName, "HiScore", "strH5", "Brad")

If ShowScores = True Then
    With frmHighScore
    .lblName1.Caption = strH1
    .lblName2.Caption = strH2
    .lblName3.Caption = strH3
    .lblName4.Caption = strH4
    .lblName5.Caption = strH5
        
    .lblScore1.Caption = h1
    .lblScore2.Caption = h2
    .lblScore3.Caption = h3
    .lblScore4.Caption = h4
    .lblScore5.Caption = h5
    .Show
    End With
    'tempStr = "High Scores:" & vbCrLf & vbCrLf & strH1 & " - " & h1 & vbCrLf & strH2 & " - " & h2 & vbCrLf & strH3 & " - " & h3 & vbCrLf & strH4 & " - " & h4 & vbCrLf & strH5 & " - " & h5
    'MsgBox tempStr, , "Hi Scores"
End If

tempStr = "High Scores:" & vbCrLf & vbCrLf & strH1 & " - " & h1 & vbCrLf & strH2 & " - " & h2 & vbCrLf & strH3 & " - " & h3 & vbCrLf & strH4 & " - " & h4 & vbCrLf & strH5 & " - " & h5

SaveHiScore = tempStr
End Function

Function ClearScore(AppName As String) As Boolean
result = MsgBox("Are you sure? You will erase ALL high scores!", vbYesNo, "Erase High Scores?")

If result = vbNo Then
        ClearScore = False
        Exit Function
End If

SaveSetting AppName, "HiScore", "strH1", "Marshall"
SaveSetting AppName, "HiScore", "strH2", "Charlotte"
SaveSetting AppName, "HiScore", "strH3", "Yogi"
SaveSetting AppName, "HiScore", "strH4", "Todd"
SaveSetting AppName, "HiScore", "strH5", "Brad"
SaveSetting AppName, "HiScore", "h1", "100"
SaveSetting AppName, "HiScore", "h2", "50"
SaveSetting AppName, "HiScore", "h3", "25"
SaveSetting AppName, "HiScore", "h4", "10"
SaveSetting AppName, "HiScore", "h5", "5"

ClearScore = True

frmHighScore.Form_Load

End Function


