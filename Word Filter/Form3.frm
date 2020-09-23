VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Word Filter"
   ClientHeight    =   2895
   ClientLeft      =   1695
   ClientTop       =   1545
   ClientWidth     =   4380
   LinkTopic       =   "Form3"
   ScaleHeight     =   2895
   ScaleWidth      =   4380
   Begin VB.OptionButton Option3 
      Caption         =   "Dont Replace Word"
      Height          =   285
      Left            =   2550
      TabIndex        =   6
      Top             =   2190
      Value           =   -1  'True
      Width           =   1755
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2940
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   2490
      Width           =   750
   End
   Begin VB.OptionButton Option2 
      Caption         =   "You Choose the Replacing CHR ->"
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Top             =   2520
      Width           =   2835
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Let Me Replace The Word"
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   2190
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   285
      Left            =   3840
      TabIndex        =   2
      Top             =   1830
      Width           =   435
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   30
      TabIndex        =   1
      Text            =   "Type hit and hit enter or Press the button!"
      Top             =   1830
      Width           =   3705
   End
   Begin VB.TextBox Text1 
      Height          =   1725
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   60
      Width           =   4245
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim DUH As String
If Option3.Value = True Then
  With Form1
      .WordFilter1.CheckWords Text2
      Text1 = Text1 & DUH & vbCrLf
  End With
End If
''
If Option1.Value = True Then
  With Form1
      .WordFilter1.CheckWords Text2, True, DUH, True
      Text1 = Text1 & DUH & vbCrLf
  End With
End If
''
If Option2.Value = True Then
  With Form1
      .WordFilter1.CheckWords Text2, True, DUH, False, True, Combo1.Text
      Text1 = Text1 & DUH & vbCrLf
  End With
End If
Text2 = ""
End Sub

Private Sub Form_Load()
Dim U As Integer
For U = 1 To 255
Combo1.AddItem Chr(U)
Next
Combo1.Text = "@"
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Dim DUH As String
''
If Option3.Value = True Then
  With Form1
      .WordFilter1.CheckWords Text2
      Text1 = Text1 & DUH & vbCrLf
  End With
End If
''
If Option1.Value = True Then
  With Form1
      .WordFilter1.CheckWords Text2, True, DUH, True
      Text1 = Text1 & DUH & vbCrLf
  End With
End If
''
If Option2.Value = True Then
  With Form1
      .WordFilter1.CheckWords Text2, True, DUH, False, True, Combo1.Text
      Text1 = Text1 & DUH & vbCrLf
  End With
End If
''
     Text2 = ""
     KeyAscii = 0
End If
End Sub
