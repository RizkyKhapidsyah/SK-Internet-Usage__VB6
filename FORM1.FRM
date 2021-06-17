VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Internet Usage Monitor"
   ClientHeight    =   3765
   ClientLeft      =   5700
   ClientTop       =   1455
   ClientWidth     =   3990
   ForeColor       =   &H00000000&
   Icon            =   "FORM1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command11 
      Caption         =   "Click Here To View Day By Day Usage"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   55
      Top             =   5160
      Width           =   3255
   End
   Begin VB.TextBox Text25 
      Height          =   285
      Left            =   720
      TabIndex        =   54
      Text            =   "Text25"
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Save Settings"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   51
      Top             =   6720
      Width           =   1260
   End
   Begin VB.TextBox Text21 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2280
      MaxLength       =   5
      TabIndex        =   49
      Text            =   "0"
      Top             =   6240
      Width           =   1260
   End
   Begin VB.TextBox Text24 
      Height          =   285
      Left            =   840
      TabIndex        =   48
      Text            =   "Text24"
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text23 
      Height          =   285
      Left            =   600
      TabIndex        =   47
      Text            =   "Text23"
      Top             =   360
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text22 
      Height          =   285
      Left            =   960
      TabIndex        =   46
      Text            =   "Text22"
      Top             =   360
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   480
      Top             =   3240
   End
   Begin VB.TextBox Text13 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2040
      MaxLength       =   10
      TabIndex        =   45
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton Command9 
      Caption         =   "ABOUT"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   44
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox Text20 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2280
      TabIndex        =   43
      Top             =   3960
      Width           =   1260
   End
   Begin VB.TextBox Text19 
      Height          =   285
      Left            =   960
      TabIndex        =   41
      Text            =   "Text19"
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text18 
      Height          =   285
      Left            =   600
      TabIndex        =   40
      Text            =   "Text18"
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text17 
      Height          =   285
      Left            =   240
      TabIndex        =   39
      Text            =   "Text17"
      Top             =   360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text16 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2280
      TabIndex        =   37
      Top             =   4320
      Width           =   1260
   End
   Begin VB.CommandButton Command8 
      Height          =   495
      Left            =   1440
      Picture         =   "FORM1.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "CLOSE DETAILS"
      Top             =   6600
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Height          =   495
      Left            =   1440
      Picture         =   "FORM1.frx":044E
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "MORE DETAILS"
      Top             =   3120
      Width           =   495
   End
   Begin VB.TextBox Text15 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2280
      MaxLength       =   5
      TabIndex        =   31
      Top             =   4680
      Width           =   1260
   End
   Begin VB.TextBox Text14 
      Height          =   285
      Left            =   960
      TabIndex        =   30
      Text            =   "Text14"
      Top             =   360
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.CommandButton Command6 
      Height          =   375
      Left            =   6000
      Picture         =   "FORM1.frx":0890
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "SWITCH BACK"
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Cost Settings"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   23
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save Cost Settings"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   19
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox Text12 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5640
      TabIndex        =   18
      Text            =   "0"
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox Text11 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5640
      TabIndex        =   17
      Text            =   "0"
      Top             =   360
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Reset Log"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Text10 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2040
      MaxLength       =   10
      TabIndex        =   13
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   600
      TabIndex        =   10
      Text            =   "Text9"
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "d"
      Height          =   495
      Left            =   480
      TabIndex        =   9
      Top             =   6600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "C"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   6600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   600
      TabIndex        =   7
      Text            =   "Text8"
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   360
      TabIndex        =   6
      Text            =   "Text7"
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   840
      TabIndex        =   5
      Text            =   "Text6"
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   480
      TabIndex        =   4
      Text            =   "Text5"
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   360
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1560
      Top             =   360
   End
   Begin VB.Shape Shape5 
      Height          =   1695
      Left            =   0
      Top             =   5880
      Width           =   3975
   End
   Begin VB.Label Label20 
      Caption         =   "min"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   53
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label Label19 
      Caption         =   "(write 0 for never)"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   52
      Top             =   6120
      Width           =   1935
   End
   Begin VB.Label Label18 
      Caption         =   "Remind Me To Close Internet After"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   50
      Top             =   5880
      Width           =   3975
   End
   Begin VB.Shape Shape4 
      Height          =   1935
      Left            =   0
      Top             =   3840
      Width           =   3975
   End
   Begin VB.Shape Shape3 
      Height          =   1575
      Left            =   0
      Top             =   2160
      Width           =   3975
   End
   Begin VB.Label Label17 
      Caption         =   "Started Logging On"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   42
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label Label16 
      Caption         =   "min"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   38
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label Label15 
      Caption         =   "Days Since Monitoring"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "www.gauravcreations.com"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      MouseIcon       =   "FORM1.frx":0CD2
      MousePointer    =   99  'Custom
      TabIndex        =   35
      Top             =   7200
      Width           =   2415
   End
   Begin VB.Label Label13 
      Caption         =   "Average Time Spent"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   4680
      Width           =   2415
   End
   Begin VB.Label Label12 
      Caption         =   "Grand Total Money Spent On Surfing"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   29
      Top             =   2760
      Width           =   3735
   End
   Begin VB.Label Label11 
      Caption         =   "ONLINE"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   2280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "OFFLINE"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "SETTINGS"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4320
      TabIndex        =   25
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "STATUS"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   0
      Width           =   975
   End
   Begin VB.Shape Shape2 
      Height          =   1815
      Left            =   4080
      Top             =   240
      Width           =   2655
   End
   Begin VB.Shape Shape1 
      Height          =   1815
      Left            =   0
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label Label7 
      Caption         =   "min"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   6360
      TabIndex        =   22
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "Impulse Cost"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4200
      TabIndex        =   21
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Impulse Time"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4200
      TabIndex        =   20
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "min"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   15
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Grand Total Time"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Disconnected"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   12
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Connected"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   360
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   1080
      Top             =   2280
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TheDate As Date   ' Declare variables.
Dim up As Integer
Dim down As Integer

Private Sub Command1_Click()
up = 1
End Sub

Private Sub Command10_Click()
Open App.Path & "\netreminder.ium" For Output As #1
Write #1, Text21.Text
Close #1
End Sub

Private Sub Command11_Click()
Form3.Show
End Sub

Private Sub Command2_Click()
down = 1
End Sub

Private Sub Command3_Click()
   Dim msg, Response   ' Declare variables.
   msg = "ARE YOU SURE YOU WANT TO RESET THE LOG"
   Response = MsgBox(msg, 48 + vbYesNo, "RESET")
   Select Case Response
      Case vbYes
      Open App.Path & "\totalog.ium" For Output As #1
      Print #1, "============================"
      Print #1, "Log Containing Day By Day Usage"
      Print #1, "                   Started On"
      Print #1, Format$(Now, "         dddd, mmm d, yyyy")
      Print #1, "============================"
      Close #1
      Open App.Path & "\netusage.ium" For Output As #1
      Write #1, 0
      Close #1
      Open App.Path & "\netdate.ium" For Output As #1
      Write #1, Text14.Text
      Close #1
      Text20.Text = Text14.Text
      Text10.Text = 0
      Text13.Text = 0
      Text15.Text = 0
      Text16.Text = 0
      Case vbNo
      Cancel = -1
    End Select
End Sub

Private Sub Command4_Click()
    Open App.Path & "\nettime.ium" For Output As #1
  
    Write #1, Text11.Text
    
    Close #1
    Open App.Path & "\netcost.ium" For Output As #1
  
    Write #1, Text12.Text
    
    Close #1

End Sub

Private Sub Command5_Click()
Form1.Width = Form1.Width + 2820
Command5.Enabled = False
End Sub

Private Sub Command6_Click()
Command5.Enabled = True
Form1.Width = Form1.Width - 2820
End Sub


Private Sub Command7_Click()
Form1.Height = Form1.Height + 3840
Command7.Enabled = False
End Sub

Private Sub Command8_Click()
Form1.Height = Form1.Height - 3840
Command7.Enabled = True
End Sub

Private Sub Command9_Click()
frmAbout.Show
End Sub

Private Sub Form_Load()
Setdefaults:
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN", "Iusage", App.Path & "\" & App.EXEName & ".exe"
Call AddToTray(Me.Icon, Me.Caption, Me)
Text14.Text = Format(Now, "dd/mm/yyyy")
Text25.Text = Format(Now, "mm/dd/yyyy") 'correct date problem in text box

Open App.Path & "\totalog.ium" For Append As #1
 
Close #1

Open App.Path & "\netreminder.ium" For Append As #1
Write #1, "0"
Close #1

Open App.Path & "\netreminder.ium" For Input As #1
Input #1, remin
Close #1
Text21.Text = remin

Open App.Path & "\netusage.ium" For Append As #1
Write #1, "0"
Close #1
Open App.Path & "\nettime.ium" For Append As #1
Write #1, "3"
Close #1
Open App.Path & "\netcost.ium" For Append As #1
Write #1, "1.2"
Close #1
    
    Open App.Path & "\nettime.ium" For Input As #1
  
    Input #1, netTime
    
    Close #1
    Text11.Text = netTime
    Open App.Path & "\netcost.ium" For Input As #1
  
    Input #1, cost
    
    Close #1
    Text12.Text = cost
    
    ' So that it doesnt append
    Open App.Path & "\nettime.ium" For Output As #1
  
    Write #1, Text11.Text
    
    Close #1
    
    Open App.Path & "\netcost.ium" For Output As #1
  
    Write #1, Text12.Text
    
    Close #1


'Check date
    Open App.Path & "\netusage.ium" For Input As #1
  
    Input #1, checkstate
    
    Close #1
If checkstate = 0 Then
    Open App.Path & "\netdate.ium" For Output As #1
  
    Write #1, Text25.Text
    
    Close #1
    Open App.Path & "\totalog.ium" For Output As #1
    Print #1, "============================"
    Print #1, "Log Containing Day By Day Usage"
    Print #1, "                   Started On"
    Print #1, Format$(Now, "         dddd, mmm d, yyyy")
    Print #1, "============================"
    Close #1
End If

'show data when form loaded
    Open App.Path & "\netusage.ium" For Input As #1
    Input #1, loadusage
    Close #1
    Text10.Text = loadusage
    Open App.Path & "\netdate.ium" For Input As #1
    Input #1, dates
    Close #1
    TheDate = dates
    datedi = DateDiff("d", TheDate, Date)
    Text15.Text = loadusage / (datedi + 1)
    totalmoney = loadusage / Val(Text11.Text) * Val(Text12.Text)
    Text13.Text = totalmoney
    Text16.Text = datedi
    connected = 0
    Open App.Path & "\netdate.ium" For Input As #1
    Input #1, logdate
    Close #1
    logdate = Format(logdate, "dd/mm/yyyy")
    Text20.Text = logdate

End Sub

Private Sub Form_Unload(Cancel As Integer)
If up = 1 And connected = 1 Then
Text4.Text = Format(Now, "hh:mm:ss AM/PM")
Text5.Text = Format(Now, "hh")
Text6.Text = Format(Now, "n")
onhro = Val(Text5.Text)
mulo = onhro * 60
addo = mulo + Val(Text6.Text)
results = addo - Add
onhr = Val(Text2.Text)
mul = onhr * 60
Add = mul + Val(Text3.Text)
results = addo - Add
    Open App.Path & "\netusage.ium" For Input As #1
  
    Input #1, prevresults
    
    Close #1
    Text8.Text = prevresults
    prevresults1 = Val(Text8.Text)
    totalresults = prevresults1 + results
    
    Open App.Path & "\netusage.ium" For Output As #1
  
    Write #1, totalresults
    
    Close #1
        Text8.Text = prevresults
    prevresults1 = Val(Text8.Text)
    totalresults = prevresults1 + results
    Text10.Text = totalresults
    Open App.Path & "\netusage.ium" For Output As #1
  
    Write #1, totalresults
    
    Close #1
     totalmoney = totalresults / Val(Text11.Text) * Val(Text12.Text)
     Text13.Text = totalmoney
    
    'average
    
    Open App.Path & "\netdate.ium" For Input As #1
  
    Input #1, dates
    
    Close #1
      TheDate = dates
      datedi = DateDiff("d", TheDate, Now)
      Text15.Text = totalresults / (datedi + 1)
      Text16.Text = datedi
 
      totalmoney1 = results / Val(Text11.Text) * Val(Text12.Text)
      Open App.Path & "\totalog.ium" For Append As #1
      Print #1, "------------------------------------------------------- "
      Print #1, "DATE                     : "; Text14.Text
      Print #1, "CONNECTED        : "; Text1.Text
      Print #1, "DISCONNECTED  : "; Text4.Text
      Print #1, "MONEY SPENT    : "; totalmoney1
      Print #1, "------------------------------------------------------- " & vbCrLf
      Close #1
 
 End If
 Call RemoveFromTray
 End Sub



Private Sub Label14_Click()
Shell "Explorer http://www.gauravcreations.com"
End Sub

Private Sub Timer1_Timer()
If IsRASConnected Then
Image1.Picture = LoadPicture(App.Path & "\lampon.bmp")
up = 1
down = 0
Else
Image1.Picture = LoadPicture(App.Path & "\lampoff.bmp")
down = 1
up = 0
End If
'check for 24:00
Text17.Text = Format(Now, "hh")
Text18.Text = Format(Now, "n")
Text19.Text = Format(Now, "s")
If Text17.Text = 23 And Text18.Text = 59 And Text19.Text = 59 Then
Unload Me
Form2.Show
End If
If up = 1 And connected = 0 Then
Command10.Enabled = False
Command5.Enabled = False
Command3.Enabled = False
Label11.Visible = True
Label10.Visible = False
Text1.Text = Format(Now, "hh:mm:ss AM/PM")
Text2.Text = Format(Now, "hh")
Text3.Text = Format(Now, "n")
Text7.Text = Format(Now, "s")
onhr = Val(Text2.Text)
mul = onhr * 60
secon = Val(Text7.Text) / 60
Add = mul + Val(Text3.Text) + secon
connected = 1
down = 0
Open App.Path & "\netreminder.ium" For Input As #1
Input #1, reminder
Close #1
Text21.Text = reminder
If reminder > 0 Then
timeleft = reminder + Add
Timer2.Enabled = True
End If
End If

If down = 1 And connected = 1 Then
Timer2.Enabled = False
Label10.Visible = True
Label11.Visible = False
Command5.Enabled = True
Command3.Enabled = True
Command10.Enabled = True
up = 0
connected = 0
Text4.Text = Format(Now, "hh:mm:ss AM/PM")
Text5.Text = Format(Now, "hh")
Text6.Text = Format(Now, "n")
Text9.Text = Format(Now, "s")
secoff = Val(Text7.Text) / 60
onhro = Val(Text5.Text)
mulo = onhro * 60
addo = mulo + Val(Text6.Text) + secoff
results = addo - Add
    onhr = Val(Text2.Text)
    mul = onhr * 60
    Add = mul + Val(Text3.Text)
    results = addo - Add
    Open App.Path & "\netusage.ium" For Input As #1
  
    Input #1, prevresults
    
    Close #1
    Text8.Text = prevresults
    prevresults1 = Val(Text8.Text)
    totalresults = prevresults1 + results
    Text10.Text = totalresults
    Open App.Path & "\netusage.ium" For Output As #1
  
    Write #1, totalresults
    
    Close #1
     totalmoney = totalresults / Val(Text11.Text) * Val(Text12.Text)
     Text13.Text = totalmoney
    
    'average
    
     Open App.Path & "\netdate.ium" For Input As #1
  
    Input #1, dates
    
    Close #1
      TheDate = dates
      datedi = DateDiff("d", TheDate, Now)
      Text15.Text = totalresults / (datedi + 1)
      Text16.Text = datedi
      
      totalmoney1 = results / Val(Text11.Text) * Val(Text12.Text)
      Open App.Path & "\totalog.ium" For Append As #1
      Print #1, "------------------------------------------------------- "
      Print #1, "DATE                     : "; Text14.Text
      Print #1, "CONNECTED        : "; Text1.Text
      Print #1, "DISCONNECTED  : "; Text4.Text
      Print #1, "MONEY SPENT     : "; totalmoney1
      Print #1, "------------------------------------------------------- " & vbCrLf
      Close #1
      
      End If
   
End Sub

Private Sub Timer2_Timer()
Text22.Text = Format(Now, "hh")
Text23.Text = Format(Now, "n")
Text24.Text = Format(Now, "s")
secofft = Val(Text24.Text) / 60
onhrot = Val(Text22.Text)
mulot = onhrot * 60
addot = mulot + Val(Text23.Text) + secofft
If addot > timeleft Then
MsgBox ("REMINDER: CLOSE YOUR INTERNET NOW"), vbInformation, "INTERNET REMINDER"
Timer2.Enabled = False
End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If RespondToTray(X) <> 0 Then Call ShowFormAgain(Me)
End Sub
