VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   ScaleHeight     =   3750
   ScaleWidth      =   5475
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtColour 
      Height          =   285
      Index           =   2
      Left            =   4800
      TabIndex        =   20
      Top             =   1020
      Width           =   615
   End
   Begin VB.TextBox txtColour 
      Height          =   285
      Index           =   1
      Left            =   4800
      TabIndex        =   19
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtColour 
      Height          =   285
      Index           =   0
      Left            =   4800
      TabIndex        =   18
      Top             =   180
      Width           =   615
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   4440
      ScaleHeight     =   315
      ScaleWidth      =   855
      TabIndex        =   17
      Top             =   1980
      Width           =   855
   End
   Begin VB.TextBox txtHTML 
      Height          =   285
      Left            =   3300
      TabIndex        =   7
      Top             =   3360
      Width           =   1995
   End
   Begin VB.TextBox txtDecimal 
      Height          =   285
      Left            =   3300
      TabIndex        =   3
      Top             =   2700
      Width           =   1995
   End
   Begin VB.TextBox txtRGB 
      Height          =   285
      Left            =   180
      TabIndex        =   2
      Top             =   3360
      Width           =   1995
   End
   Begin VB.TextBox txtHex 
      Height          =   285
      Left            =   180
      TabIndex        =   1
      Top             =   2700
      Width           =   1995
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   180
      ScaleHeight     =   495
      ScaleWidth      =   3975
      TabIndex        =   0
      Top             =   1800
      Width           =   3975
   End
   Begin ComctlLib.Slider sldRed 
      Height          =   255
      Left            =   900
      TabIndex        =   9
      Top             =   240
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   450
      _Version        =   327682
      Max             =   255
      SelStart        =   1
      TickStyle       =   3
      Value           =   1
   End
   Begin ComctlLib.Slider sldGreen 
      Height          =   255
      Left            =   900
      TabIndex        =   10
      Top             =   660
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   450
      _Version        =   327682
      Max             =   255
      SelStart        =   1
      TickStyle       =   3
      Value           =   1
   End
   Begin ComctlLib.Slider sldBlue 
      Height          =   255
      Left            =   900
      TabIndex        =   11
      Top             =   1080
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   450
      _Version        =   327682
      Max             =   255
      SelStart        =   1
      TickStyle       =   3
      Value           =   1
   End
   Begin VB.Label Label9 
      Caption         =   "Result"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   1500
      Width           =   615
   End
   Begin VB.Label Label8 
      Caption         =   "(Opposite)"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   15
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Blue"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1080
      Width           =   675
   End
   Begin VB.Label Label6 
      Caption         =   "Green"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   660
      Width           =   675
   End
   Begin VB.Label Label5 
      Caption         =   "Red"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   240
      Width           =   675
   End
   Begin VB.Label Label4 
      Caption         =   "HTML"
      Height          =   255
      Left            =   3360
      TabIndex        =   8
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Decimal"
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   2460
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "RGB"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   3120
      Width           =   1875
   End
   Begin VB.Label Label1 
      Caption         =   "Hex"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   2460
      Width           =   1875
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const CurrentModule As String = "Form1"

Dim cColours As clsColours

Dim iRed, iGreen, iBlue As Long


Private Sub Form_Load()

On Error GoTo Err_Form_Load

    Set cColours = New clsColours

    iRed = 80
    iGreen = 160
    iBlue = 240

    sldRed.Value = iRed
    sldGreen.Value = iGreen
    sldBlue.Value = iBlue

    ColourChanged

Exit Sub

Err_Form_Load:
    Screen.MousePointer = vbDefault
    HandleError CurrentModule, "Form_Load", Err.Number, Err.Description

End Sub

Private Sub sldRed_Scroll()

On Error GoTo Err_sldRed_Scroll

    iRed = sldRed.Value
    ColourChanged

Exit Sub

Err_sldRed_Scroll:
    Screen.MousePointer = vbDefault
    HandleError CurrentModule, "sldRed_Scroll", Err.Number, Err.Description

End Sub

Private Sub sldGreen_Scroll()

On Error GoTo Err_sldGreen_Scroll

    iGreen = sldGreen.Value
    ColourChanged

Exit Sub

Err_sldGreen_Scroll:
    Screen.MousePointer = vbDefault
    HandleError CurrentModule, "sldGreen_Scroll", Err.Number, Err.Description

End Sub

Private Sub sldBlue_Scroll()

On Error GoTo Err_sldBlue_Scroll

    iBlue = sldBlue.Value
    ColourChanged

Exit Sub

Err_sldBlue_Scroll:
    Screen.MousePointer = vbDefault
    HandleError CurrentModule, "sldBlue_Scroll", Err.Number, Err.Description

End Sub

Private Sub ColourChanged()

On Error GoTo Err_ColourChanged

    Picture1.BackColor = RGB(iRed, iGreen, iBlue)
    Picture2.BackColor = cColours.GetOppositeColour(iRed, iGreen, iBlue)

    txtHex = cColours.GetHexFromRGB(iRed, iGreen, iBlue)

    txtRGB = "RGB(" & iRed & ", " & iGreen & ", " & iBlue & ")"

    txtDecimal = cColours.GetDecimalFromRGB(iRed, iGreen, iBlue)

    txtHTML = cColours.GetHTMLColourCodeFromRGB(iRed, iGreen, iBlue)

    txtColour(0).Text = iRed
    txtColour(1).Text = iGreen
    txtColour(2).Text = iBlue

Exit Sub

Err_ColourChanged:
    Screen.MousePointer = vbDefault
    HandleError CurrentModule, "ColourChanged", Err.Number, Err.Description

End Sub

Private Sub txtColour_Change(Index As Integer)

On Error Resume Next

    If txtColour(Index).Text = "" Then
        txtColour(Index).Text = 0
    End If

    If CLng(txtColour(Index).Text) > 255 Then
        txtColour(Index).Text = 255
        Highlight txtColour(Index)
    End If

    If Index = 0 Then
        iRed = CLng(txtColour(Index).Text)
        sldRed.Value = iRed
    ElseIf Index = 1 Then
        iGreen = CLng(txtColour(Index).Text)
        sldGreen.Value = iGreen
    ElseIf Index = 2 Then
        iBlue = CLng(txtColour(Index).Text)
        sldBlue.Value = iBlue
    End If

    ColourChanged

End Sub

Private Sub txtColour_GotFocus(Index As Integer)

On Error GoTo Err_txtColour_GotFocus

    Highlight txtColour(Index)

Exit Sub

Err_txtColour_GotFocus:
    Screen.MousePointer = vbDefault
    HandleError CurrentModule, "txtColour_GotFocus", Err.Number, Err.Description

End Sub

Private Sub txtColour_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Err_txtColour_KeyPress

    If KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then Exit Sub

    If Not IsNumeric(Chr(KeyAscii)) Then
        KeyAscii = 0
    End If

Exit Sub

Err_txtColour_KeyPress:
    Screen.MousePointer = vbDefault
    HandleError CurrentModule, "txtColour_KeyPress", Err.Number, Err.Description

End Sub

Private Sub Highlight(obj As Object)

On Error Resume Next

    With obj
        .SetFocus
        .SelStart = 0
        .SelLength = Len(obj)
    End With

End Sub

Private Sub txtDecimal_GotFocus()

On Error GoTo Err_txtDecimal_GotFocus

    Highlight txtDecimal

Exit Sub

Err_txtDecimal_GotFocus:
    Screen.MousePointer = vbDefault
    HandleError CurrentModule, "txtDecimal_GotFocus", Err.Number, Err.Description

End Sub

Private Sub txtHex_GotFocus()

On Error GoTo Err_txtHex_GotFocus

    Highlight txtHex

Exit Sub

Err_txtHex_GotFocus:
    Screen.MousePointer = vbDefault
    HandleError CurrentModule, "txtHex_GotFocus", Err.Number, Err.Description

End Sub

Private Sub txtHTML_GotFocus()

On Error GoTo Err_txtHTML_GotFocus

    Highlight txtHTML

Exit Sub

Err_txtHTML_GotFocus:
    Screen.MousePointer = vbDefault
    HandleError CurrentModule, "txtHTML_GotFocus", Err.Number, Err.Description

End Sub

Private Sub txtRGB_GotFocus()

On Error GoTo Err_txtRGB_GotFocus

    Highlight txtRGB

Exit Sub

Err_txtRGB_GotFocus:
    Screen.MousePointer = vbDefault
    HandleError CurrentModule, "txtRGB_GotFocus", Err.Number, Err.Description

End Sub
