VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Internet Usage Monitor"
   ClientHeight    =   1830
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5085
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1263.098
   ScaleMode       =   0  'User
   ScaleWidth      =   4775.078
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   1920
      TabIndex        =   0
      Top             =   1320
      Width           =   1260
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmAbout.frx":000C
      Stretch         =   -1  'True
      Top             =   120
      Width           =   585
   End
   Begin VB.Label lblTitle 
      Caption         =   "INTERNET USAGE MONITOR "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   840
      TabIndex        =   1
      Top             =   240
      Width           =   3885
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version 1.0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   840
      TabIndex        =   2
      Top             =   720
      Width           =   3885
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub lblDisclaimer_Click()
Shell "Explorer https://www.google.com"
End Sub


