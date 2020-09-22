VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "comct232.ocx"
Begin VB.UserControl TimeBox 
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1350
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   ScaleHeight     =   360
   ScaleWidth      =   1350
   Begin VB.PictureBox PicFrame 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   1335
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   1335
      Begin VB.PictureBox picTime 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   260
         Left            =   60
         ScaleHeight     =   255
         ScaleWidth      =   930
         TabIndex        =   8
         Top             =   45
         Width           =   925
         Begin VB.TextBox txtHN 
            BorderStyle     =   0  'None
            Height          =   255
            Index           =   2
            Left            =   680
            MaxLength       =   2
            TabIndex        =   2
            Top             =   0
            Width           =   225
         End
         Begin VB.TextBox txtHN 
            BorderStyle     =   0  'None
            Height          =   255
            Index           =   0
            Left            =   0
            MaxLength       =   2
            TabIndex        =   0
            Top             =   0
            Width           =   225
         End
         Begin VB.TextBox txtHN 
            BorderStyle     =   0  'None
            Height          =   255
            Index           =   1
            Left            =   340
            MaxLength       =   2
            TabIndex        =   1
            Top             =   0
            Width           =   225
         End
         Begin VB.Label lblColon 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
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
            Index           =   1
            Left            =   580
            TabIndex        =   12
            Top             =   0
            Width           =   45
         End
         Begin VB.Label lblColon 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
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
            Index           =   0
            Left            =   240
            TabIndex        =   9
            Top             =   0
            Width           =   45
         End
      End
      Begin VB.TextBox txtAMPM 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   625
         TabIndex        =   5
         Text            =   "AM"
         Top             =   495
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox txtMinutes 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   390
         TabIndex        =   4
         Text            =   "01"
         Top             =   480
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.TextBox txtColon 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   315
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   ":"
         Top             =   495
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.TextBox txtHours 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   3
         Text            =   "01"
         Top             =   495
         Visible         =   0   'False
         Width           =   225
      End
      Begin ComCtl2.UpDown uHN 
         Height          =   330
         Left            =   1020
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   -10
         WhatsThisHelpID =   200
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   582
         _Version        =   327681
         Max             =   60
         Min             =   -1
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtBackTime 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   0
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   0
         Width           =   1020
      End
   End
End
Attribute VB_Name = "TimeBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Author: Carlos Alberto S.
'Date: 19Jan2007
'Purpose: emulate the MS Windows textbox time input control. It can show or hide the second box.
'Note 1: this control was made to use in another project; it has the options that I needed.
'       I don't have time enough to create a full featured control but I'm sure you can
'       easily change it to fit your needs too.
'Note 2: the control uses the MS Up/Down control; some people don't like it since you have to
'        include an OCX; anyway, you can easily change for another user control or even
'         the VScrollBar
'Based on the following code:
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=6359&lngWId=1
'There are some codes here from other PSC users too. Thank you all!
'
'Update #1 (20/Jan/2007): fixed leading zero issue
'Update #2 (23/Jan/2007): fine code adjustments thanks to Roger G.
'
Option Explicit
Private m_DescriptionFormat As String
'Detect if TAB key was pressed
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Const VK_TAB = &H9
Private lngTime As Long
Private isWorking As Boolean
Private isN As Boolean
Private isSel As Boolean
Private isSec As Boolean
Private strTimeSelected As String
Event Change(iTimeText As Integer, vVal As Variant)

Private Function isTabPress() As Boolean

    Dim iRetVal As Integer
    iRetVal = GetKeyState(VK_TAB)
    If iRetVal = -128 Or iRetVal = -127 Then isTabPress = True

End Function

Private Sub txtBackTime_GotFocus()
    txtHN(0).SetFocus
End Sub

Private Sub txtHN_Change(Index As Integer)

    On Error Resume Next

    isWorking = True

    txtHN(Index) = NumberOrNoNumber(txtHN(Index), True)
    If txtHN(Index) = "" Then txtHN(Index) = "00"
    'If txtHN(Index) = "0" Then txtHN(Index) = "00"

    Select Case Index
    Case 0
        If txtHN(0) >= 24 Then txtHN(0) = "00"
        If lngTime = 0 Then If txtHN(0) <> "" Then uHN.Value = Val(txtHN(0))
    Case 1
        If txtHN(1) >= 60 Then txtHN(1) = "00"    ': txtHN(0) = Val(txtHN(0)) + 1
        If lngTime = 1 Then If txtHN(1) <> "" Then uHN.Value = Val(txtHN(1))
    Case 2
        If txtHN(2) >= 60 Then txtHN(2) = "00"    ': txtHN(1) = Val(txtHN(1)) + 1
        If lngTime = 2 Then If txtHN(2) <> "" Then uHN.Value = Val(txtHN(2))
    End Select

    m_DescriptionFormat = txtHN(0) & ":" & txtHN(1) & ":" & txtHN(2)

    RaiseEvent Change(Index, txtHN(Index))

    DoEvents
    isWorking = False

    Err.Clear
    On Error GoTo 0

End Sub

Private Sub txtHN_GotFocus(Index As Integer)

    On Error Resume Next

    lngTime = Index

    Select Case Index
    Case 0
        If txtHN(0) <> "" Then uHN.Value = Val(txtHN(0)): uHN.Max = 24
    Case 1
        If txtHN(1) <> "" Then uHN.Value = Val(txtHN(1)): uHN.Max = 60
    Case 2
        If txtHN(2) <> "" Then uHN.Value = Val(txtHN(2)): uHN.Max = 60
    End Select

    If isSel Then
        txtHN(Index).SelStart = 0
        txtHN(Index).SelLength = Len(txtHN(Index))
    Else
        txtHN(Index).SelStart = 0
    End If

    Err.Clear
    On Error GoTo 0

End Sub

Private Sub txtHN_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    Dim lngTmp As Long

    Select Case KeyCode
    Case 38
        txtHN(Index) = Val(txtHN(Index)) + 1
    Case 40
        'negative value is not handled in txtHN change event due NumberOrNoNumber
        lngTmp = Val(txtHN(Index)) - 1
        If lngTmp = -1 Then
            If lngTime = 0 Then
                txtHN(Index) = 23
            Else
                txtHN(Index) = 59
            End If
        Else
            txtHN(Index) = Val(txtHN(Index)) - 1
        End If
    End Select

End Sub

Private Sub txtHN_LostFocus(Index As Integer)

    Select Case Index
    Case 0
        txtHN(0) = Format(txtHN(0), "00")
        If isTabPress Then txtHN(1).SetFocus
    Case 1
        txtHN(1) = Format(txtHN(1), "00")
        If isSec Then If isTabPress Then txtHN(2).SetFocus
    Case 2
        txtHN(2) = Format(txtHN(2), "00")
        'let tab go outside the control!
    End Select

End Sub

Private Sub uHN_Change()
    If isWorking Then Exit Sub
    If uHN.Value = -1 Then
        Select Case lngTime
        Case 0
            txtHN(0) = 23
        Case 1
            txtHN(1) = 59
        Case 2
            txtHN(2) = 59
        End Select
    End If
    txtHN(lngTime) = Format(uHN.Value, "00")
End Sub

Private Function NumberOrNoNumber(ByVal StrToCheck As String, ByVal Numbers As Boolean, Optional NumericTextTarget As TextBox)

    Dim Nstr As String
    Dim Tstr As String
    Dim I As Integer

    For I = 1 To Len(StrToCheck)
        If IsNumeric(Mid(StrToCheck, I, 1)) Then
            Nstr = Nstr & Mid(StrToCheck, I, 1)
        Else
            Tstr = Tstr & Mid(StrToCheck, I, 1)
        End If
    Next

    If Numbers Then NumberOrNoNumber = Nstr Else NumberOrNoNumber = Tstr


    On Error Resume Next
    NumericTextTarget = Nstr
    On Error GoTo 0

End Function

Private Sub UserControl_InitProperties()

    txtHN(0).Text = UserControl.Extender.Name
    txtHN(1).Text = UserControl.Extender.Name
    txtHN(2).Text = UserControl.Extender.Name
    'm_DescriptionFormat = "00:00:00"
    m_DescriptionFormat = Time
    lngTime = 0
    If txtHN(0) <> "" Then uHN.Value = Val(txtHN(0)): uHN.Max = 24
    isSec = True
    txtBackTime.Width = 1020
    uHN.Left = 1020
    picTime.Width = 925
    UserControl.Height = txtBackTime.Height
    UserControl.Width = 1260
    lblColon(1).Visible = True
    txtHN(2).Visible = True
    txtHN(0).FontName = UserControl.FontName
    txtHN(1).FontName = UserControl.FontName
    txtHN(2).FontName = UserControl.FontName
    lblColon(0).FontName = UserControl.FontName
    lblColon(1).FontName = UserControl.FontName

End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = 0
End Sub
Private Sub UserControl_Resize()

    If UserControl.Width > 1000 Then
        txtBackTime.Width = 1020
        uHN.Left = 1020
        picTime.Width = 925
        UserControl.Height = txtBackTime.Height
        UserControl.Width = 1260
        lblColon(1).Visible = True
        txtHN(2).Visible = True
        isSec = True
    Else
        txtBackTime.Width = 735
        uHN.Left = 735
        picTime.Width = 615
        UserControl.Height = txtBackTime.Height
        UserControl.Width = 970
        lblColon(1).Visible = False
        txtHN(2).Visible = False
        isSec = False
    End If

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    isSel = PropBag.ReadProperty("Seltext", False)
    ShowSecond = PropBag.ReadProperty("ShowSecond", False)
    Text = PropBag.ReadProperty("Text", Time)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error Resume Next
    PropBag.WriteProperty "Seltext", isSel, False
    PropBag.WriteProperty "ShowSecond", ShowSecond, False
    PropBag.WriteProperty "Text", Text, Time
    PropBag.WriteProperty "Font", UserControl.Font, Ambient.Font
End Sub
Property Get ShowSecond() As Boolean
    ShowSecond = isSec
End Property
Property Let ShowSecond(NewVal As Boolean)
    isSec = NewVal
    PropertyChanged "ShowSecond"
    If isSec Then
        txtBackTime.Width = 1020
        uHN.Left = 1020
        picTime.Width = 925
        UserControl.Height = txtBackTime.Height
        UserControl.Width = 1260
        lblColon(1).Visible = True
        txtHN(2).Visible = True
        isSec = True
        txtHN(0).SetFocus
    Else
        txtBackTime.Width = 735
        uHN.Left = 735
        picTime.Width = 615
        UserControl.Height = txtBackTime.Height
        UserControl.Width = 970
        lblColon(1).Visible = False
        txtHN(2).Visible = False
        isSec = False
        txtHN(0).SetFocus
    End If
End Property

Property Get Seltext() As Boolean
    Seltext = isSel
End Property
Property Let Seltext(NewVal As Boolean)
    isSel = NewVal
    PropertyChanged "Seltext"
End Property

Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal fVal As Font)

    Set UserControl.Font = fVal
    Set txtHN(0).Font = fVal
    Set txtHN(1).Font = fVal
    Set txtHN(2).Font = fVal
    Set lblColon(0).Font = fVal
    Set lblColon(1).Font = fVal
    PropertyChanged "Font"

End Property

Public Property Get Text() As String
    Text = m_DescriptionFormat
End Property

Public Property Let Text(sVal As String)
    m_DescriptionFormat = sVal
    PropertyChanged "Text"
    txtHN(0) = Left(sVal, 2)
    txtHN(1) = Mid(sVal, 4, 2)
    txtHN(2) = Right(sVal, 2)
End Property
