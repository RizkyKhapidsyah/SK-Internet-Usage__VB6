VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Day By Day Usage"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3180
   Icon            =   "dayby.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   3180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "EXIT"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   3360
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   3135
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
Form1.Show
End Sub

Private Sub Form_Load()
Open App.Path & "\totalog.ium" For Input As #1
Do While Not EOF(1)   ' Check for end of file.
  linesfromfile = StrConv(InputB(LOF(1), 1), vbUnicode)
  Text1.Text = linesfromfile
Loop
Close #1
End Sub
