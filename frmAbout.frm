VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3540
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   216
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   236
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picCont 
      BorderStyle     =   0  'None
      Height          =   1560
      Left            =   150
      ScaleHeight     =   104
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   217
      TabIndex        =   4
      Top             =   975
      Width           =   3255
      Begin VB.PictureBox picCredits 
         BorderStyle     =   0  'None
         Height          =   5280
         Left            =   225
         ScaleHeight     =   352
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   182
         TabIndex        =   5
         Top             =   -2595
         Width           =   2730
         Begin VB.Label Label8 
            Caption         =   "VB IDE           Terry O'Brien"
            Height          =   375
            Left            =   870
            TabIndex        =   10
            Top             =   3000
            Width           =   1080
         End
         Begin VB.Label Label7 
            Caption         =   "Moral Support  My Parents"
            Height          =   435
            Left            =   870
            TabIndex        =   9
            Top             =   2220
            Width           =   1170
         End
         Begin VB.Label Label6 
            Caption         =   "Art                       Marshall Hecht"
            Height          =   525
            Left            =   885
            TabIndex        =   8
            Top             =   1470
            Width           =   1185
         End
         Begin VB.Label Label5 
            Caption         =   "Programmer      Marshall Hecht"
            Height          =   405
            Left            =   900
            TabIndex        =   7
            Top             =   810
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Director           Marshall Hecht"
            Height          =   465
            Left            =   900
            TabIndex        =   6
            Top             =   90
            Width           =   1080
         End
      End
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "OK"
      Height          =   375
      Left            =   2535
      TabIndex        =   2
      Top             =   2745
      Width           =   885
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   9
      X2              =   9
      Y1              =   64
      Y2              =   170
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   227
      X2              =   227
      Y1              =   168
      Y2              =   64
   End
   Begin VB.Image Image3 
      Height          =   600
      Left            =   375
      Picture         =   "frmAbout.frx":0000
      Top             =   2610
      Width           =   1500
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "by"
      Height          =   195
      Left            =   165
      TabIndex        =   3
      Top             =   2835
      Width           =   165
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   10
      X2              =   229
      Y1              =   169
      Y2              =   169
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   8
      X2              =   228
      Y1              =   64
      Y2              =   64
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Version 1.0"
      Height          =   195
      Left            =   885
      TabIndex        =   1
      Top             =   675
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ocean Reef"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   840
      TabIndex        =   0
      Top             =   165
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   180
      Picture         =   "frmAbout.frx":0364
      Top             =   150
      Width           =   480
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdReturn_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    ScrollCredits
End Sub

Private Sub Form_Load()
    picCredits.Left = 18
    picCredits.Top = 100
                
End Sub

Sub ScrollCredits()
    
    Do
        If tTime2 + ms_Delay2 <= GetTickCount() Then
        'Reset the timer
        tTime2 = GetTickCount()
        picCredits.Top = picCredits.Top - 1
        
        DoEvents 'If you don't put this in, Windows screws up
                 'It returns control to Windows to handle prg
                 'requests
        End If
    Loop Until picCredits.Top > picCredits.Top + picCredits.Height
End Sub


Private Sub Image1_Click()
    MsgBox "*****WARNING******" & vbCrLf & vbCrLf & "If you are a terrorist, may" & vbCrLf & "bombs rain down on your house" & vbCrLf & "and viruses format your" & vbCrLf & "harddrive!", vbCritical, "WARNING"
End Sub

Private Sub Image3_Click()
    MsgBox "Guess what! You found the Easter Egg!" _
    & vbCrLf & vbCrLf & "To get infinite health:" _
    & vbCrLf & "  Get the First Aid kit" & vbCrLf _
    & "  Go left off the screen until" & vbCrLf & _
    "  your life bar shoots up." & vbCrLf & vbCrLf _
    & "Don't tell anyone else!", vbInformation, "Easter Egg"
End Sub
