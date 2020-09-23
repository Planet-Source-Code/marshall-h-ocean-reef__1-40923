VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3570
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   181
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   238
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdReturn 
      Caption         =   "OK"
      Height          =   390
      Left            =   2505
      TabIndex        =   1
      Top             =   2265
      Width           =   1005
   End
   Begin VB.TextBox Text1 
      Height          =   2115
      Left            =   75
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmHelp.frx":0000
      Top             =   75
      Width           =   3435
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdReturn_Click()
    Me.Hide
End Sub
