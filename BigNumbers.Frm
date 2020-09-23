VERSION 4.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Big Numbers"
   ClientHeight    =   5760
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   9300
   Height          =   6165
   Icon            =   "BigNumbers.frx":0000
   Left            =   1080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9300
   Top             =   1170
   Width           =   9420
   Begin VB.CommandButton Command27 
      Caption         =   "EXIT"
      Height          =   300
      Left            =   8220
      TabIndex        =   20
      Top             =   5400
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   18
      Top             =   4680
      Width           =   8775
   End
   Begin VB.CommandButton Command25 
      Caption         =   "ÖA"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Symbol"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7560
      TabIndex        =   17
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton Command24 
      Caption         =   "A^B"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6480
      TabIndex        =   16
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton Command23 
      Caption         =   "A!"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5400
      TabIndex        =   15
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton Command21 
      Caption         =   "To A"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3240
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command20 
      Caption         =   "To B"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4560
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command19 
      Caption         =   "ß"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Symbol"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2520
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command18 
      Caption         =   "ßÝ"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Symbol"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3840
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Ý"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Symbol"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5160
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command16 
      Caption         =   "A/B"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4320
      TabIndex        =   6
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton Command12 
      Caption         =   "A-B"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2160
      TabIndex        =   5
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "A*B"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3240
      TabIndex        =   4
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "A+B"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1020
      TabIndex        =   3
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   1575
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   3000
      Width           =   8775
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "BigNumbers.frx":030A
      Top             =   1320
      Width           =   8775
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "BigNumbers.frx":030E
      Top             =   120
      Width           =   8775
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "R:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   19
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "="
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   14
      Top             =   3120
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "B:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   13
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "A:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   12
      Top             =   360
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit




Private Sub Command12_Click()
Dim a(MaxLength) As Long
Dim b(MaxLength) As Long
Dim c(MaxLength) As Long
Dim LengthA As Long
Dim LengthB As Long
Dim LengthC As Long
MousePointer = 11

Call TextToBigNumber(Text1.Text, a, LengthA)
Call TextToBigNumber(Text2.Text, b, LengthB)


Call MinusBSigned(a, LengthA, b, LengthB, c, LengthC)
Text3.Text = BigNumberToText(c, LengthC)
Text4.Text = ""
MousePointer = 0

End Sub




Private Sub Command16_Click()
Dim a(MaxLength) As Long
Dim b(MaxLength) As Long
Dim c(MaxLength) As Long
Dim d(MaxLength) As Long
Dim LengthA As Long
Dim LengthB As Long
Dim LengthC As Long
Dim LengthD As Long
MousePointer = 11

Call TextToBigNumber(Text1.Text, a, LengthA)
Call TextToBigNumber(Text2.Text, b, LengthB)
Call DivBSigned(a, LengthA, b, LengthB, c, LengthC, d, LengthD)
Text3.Text = BigNumberToText(c, LengthC)
Text4.Text = BigNumberToText(d, LengthD)
MousePointer = 0

End Sub


Private Sub Command17_Click()
Text1.Text = Text2.Text

End Sub

Private Sub Command18_Click()
Dim pom As String
pom = Text1.Text
Text1.Text = Text2.Text
Text2.Text = pom

End Sub


Private Sub Command19_Click()
Text2.Text = Text1.Text
End Sub


Private Sub Command20_Click()
Text2.Text = Text3.Text

End Sub

Private Sub Command21_Click()
Text1.Text = Text3.Text

End Sub


Private Sub Command23_Click()
Dim c(MaxLength) As Long
Dim LengthC As Long
MousePointer = 11
If IsItBigNumber(Text1.Text) Then
    If Abs(Val(Text1.Text)) <= 32767 Then
        Call Factorial(Val(Text1.Text), c, LengthC)
        Text3.Text = BigNumberToText(c, LengthC)
    End If
Else
    Text3.Text = "0"
End If
Text4.Text = ""
MousePointer = 0

End Sub

Private Sub Command24_Click()
Dim a(MaxLength) As Long
Dim LengthA As Long
Dim c(MaxLength) As Long
Dim LengthC As Long
MousePointer = 11
Call TextToBigNumber(Text1.Text, a, LengthA)
If IsItBigNumber(Text2.Text) Then
    If Abs(Val(Text2.Text)) <= 32767 Then
        Call PowerB(a, LengthA, Val(Text2.Text), c, LengthC)
        Text3.Text = BigNumberToText(c, LengthC)
    End If
Else
    Text3.Text = "0"
End If
Text4.Text = ""
MousePointer = 0

End Sub

Private Sub Command25_Click()
Dim a(MaxLength) As Long
Dim c(MaxLength) As Long
Dim d(MaxLength) As Long
Dim LengthA As Long
Dim LengthC As Long
Dim LengthD As Long
MousePointer = 11

Call TextToBigNumber(Text1.Text, a, LengthA)
Call SqrtB(a, LengthA, c, LengthC, d, LengthD)
Text3.Text = BigNumberToText(c, LengthC)
Text4.Text = BigNumberToText(d, LengthD)
MousePointer = 0

End Sub


Private Sub Command27_Click()
Unload Me
End
End Sub

Private Sub Command3_Click()
Dim a(MaxLength) As Long
Dim b(MaxLength) As Long
Dim c(MaxLength) As Long
Dim LengthA As Long
Dim LengthB As Long
Dim LengthC As Long
MousePointer = 11

Call TextToBigNumber(Text1.Text, a, LengthA)
Call TextToBigNumber(Text2.Text, b, LengthB)
Call AddBSigned(a, LengthA, b, LengthB, c, LengthC)
Text3.Text = BigNumberToText(c, LengthC)
Text4.Text = ""
MousePointer = 0
End Sub

Private Sub Command4_Click()
Dim a(MaxLength) As Long
Dim b(MaxLength) As Long
Dim c(MaxLength) As Long
Dim LengthA As Long
Dim LengthB As Long
Dim LengthC As Long
MousePointer = 11

Call TextToBigNumber(Text1.Text, a, LengthA)
Call TextToBigNumber(Text2.Text, b, LengthB)
Call MultBSigned(a, LengthA, b, LengthB, c, LengthC)
Text3.Text = BigNumberToText(c, LengthC)
Text4.Text = ""
MousePointer = 0

End Sub





Private Sub Form_Load()
Text1.Text = "1"
Text2.Text = "1"
Text3.Text = "0"
Text4.Text = "0"
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case Asc("-"), Asc("+"), Asc("0") To Asc("9")
Case Else
    KeyAscii = 0
End Select

End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case Asc("-"), Asc("+"), Asc("0") To Asc("9")
Case Else
    KeyAscii = 0
End Select

End Sub


Private Sub Text3_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case Asc("-"), Asc("+"), Asc("0") To Asc("9")
Case Else
    KeyAscii = 0
End Select

End Sub


Private Sub Text4_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case Asc("-"), Asc("+"), Asc("0") To Asc("9")
Case Else
    KeyAscii = 0
End Select

End Sub


