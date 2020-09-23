VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Word Filter OCX"
   ClientHeight    =   4215
   ClientLeft      =   1695
   ClientTop       =   1545
   ClientWidth     =   3900
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   3900
   Begin Project1.WordFilter WordFilter1 
      Height          =   345
      Left            =   1710
      TabIndex        =   8
      Top             =   4560
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   609
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   4020
   End
   Begin VB.Frame Frame1 
      Caption         =   "Word Filter"
      Height          =   3855
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   3555
      Begin VB.CommandButton Command5 
         Caption         =   "Try Me!"
         Height          =   285
         Left            =   210
         TabIndex        =   10
         Top             =   3420
         Width           =   2925
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Find"
         Height          =   315
         Left            =   210
         TabIndex        =   9
         Top             =   1710
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Edit"
         Height          =   315
         Left            =   210
         TabIndex        =   5
         Top             =   2130
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Remove"
         Height          =   315
         Left            =   210
         TabIndex        =   4
         Top             =   2580
         Width           =   975
      End
      Begin VB.ListBox List1 
         Height          =   2595
         Left            =   1290
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   330
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1260
         TabIndex        =   2
         Top             =   3060
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ADD"
         Height          =   315
         Left            =   210
         TabIndex        =   1
         Top             =   3030
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   225
         Left            =   450
         TabIndex        =   7
         Top             =   570
         Width           =   315
      End
      Begin VB.Label Label1 
         Caption         =   "Word Count"
         Height          =   255
         Left            =   180
         TabIndex        =   6
         Top             =   330
         Width           =   885
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim S As Integer

Private Sub Command1_Click()
Dim REDF As Boolean
REDF = WordFilter1.AddWord(Text1.Text)
If REDF Then
   List1.AddItem Text1.Text
End If

End Sub

Private Sub Command2_Click()
WordFilter1.RemoveWord List1.Text, True, List1
End Sub

Private Sub Command3_Click()
Dim UGh As String
UGh = InputBox("Edit This Word to what?", "Word Filter Editor!", List1.Text)
If UGh = "" Then Exit Sub
WordFilter1.RemoveWord UGh, True, List1
WordFilter1.AddWord UGh
End Sub

Private Sub Command4_Click()
Form2.Show
Form2.Top = Form1.Top
Form2.Left = Form1.Left + Form1.Width
Form2.Text1.SetFocus
End Sub

Private Sub Command5_Click()
Form3.Show
Form3.Text2.SetFocus
End Sub

Private Sub Form_Load()
WordFilter1.ListLoad App.Path & "\BadWordList.txt", List1
End Sub

Private Sub Timer1_Timer()
Label2.Caption = List1.ListCount
If Len(List1.Text) < 1 Then
   Command3.Enabled = False
 ElseIf Len(List1.Text) > 0 Then
   Command3.Enabled = True
End If
End Sub



