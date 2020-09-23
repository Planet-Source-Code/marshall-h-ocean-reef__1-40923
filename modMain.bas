Attribute VB_Name = "modMain"
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public tTime As Long
Public Const ms_Delay = 1

Public tTime2 As Long
Public Const ms_Delay2 = 40

Public GameOn As Boolean

Public Score As Long
Public Health As Integer
Public Level As Integer

Public DiverStage As Integer
Public DiverAction As Boolean
Public DiverDir As String
Public DiverMoving As Boolean
Public SwimAmount

Public HaveKey As Boolean
Public HaveFlippers As Boolean
Public AtBoss As Boolean

Public FishSpeed
Public BossSpeed
Public BossDir As String

Public GroundY As Integer

Public Sub PlayWAV(sFileName As String, Optional Synchronous As Boolean = False)
    If Synchronous = True Then
        sndPlaySound App.Path & "\" & sFileName, 0
    Else
        sndPlaySound App.Path & "\" & sFileName, 1
    End If
End Sub
