VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Time input control"
   ClientHeight    =   2085
   ClientLeft      =   105
   ClientTop       =   405
   ClientWidth     =   3195
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   3195
   StartUpPosition =   3  'Windows Default
   Begin Project1.TimeBox TimeBox1 
      Height          =   315
      Left            =   240
      TabIndex        =   5
      Top             =   360
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      ShowSecond      =   -1  'True
      Text            =   "21:44:48"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox chkOpt 
      Caption         =   "ShowSecond"
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   2
      Top             =   600
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CheckBox chkOpt 
      Caption         =   "Seltext"
      Height          =   255
      Index           =   0
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get text"
      Height          =   330
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "You can use arrow up and down to change the value"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   2775
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblTxt 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1800
      TabIndex        =   3
      Top             =   1120
      Width           =   45
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkOpt_Click(Index As Integer)

    TimeBox1.Seltext = CBool(chkOpt(0))
    TimeBox1.ShowSecond = CBool(chkOpt(1))

End Sub

Private Sub Command1_Click()
    lblTxt = TimeBox1.Text
End Sub


Private Sub Form_Load()

    'initial control value = current time
    TimeBox1.Text = Time
    'but you can set a custom one e.g. from an ini file:
    'TimeBox1.Text = "00:20:40"

End Sub

Private Sub TimeBox1_Change(iTimeText As Integer, vVal As Variant)

    Select Case iTimeText
    Case 0
        Debug.Print "Hour changed " & vVal
    Case 1
        Debug.Print "Minute changed " & vVal
    Case 2
        Debug.Print "Second changed " & vVal
    End Select

    Debug.Print TimeBox1.Text

End Sub
