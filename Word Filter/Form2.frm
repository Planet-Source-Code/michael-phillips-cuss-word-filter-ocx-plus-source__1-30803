VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Word Filter"
   ClientHeight    =   1320
   ClientLeft      =   6150
   ClientTop       =   1650
   ClientWidth     =   3075
   LinkTopic       =   "Form2"
   ScaleHeight     =   1320
   ScaleWidth      =   3075
   Begin VB.Frame Frame1 
      Caption         =   "Find Word"
      Height          =   1275
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   3045
      Begin VB.CommandButton Command1 
         Caption         =   "Find"
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   810
         Width           =   2745
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   450
         Width           =   2745
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
With Form1
 .List1.ListIndex = .WordFilter1.Findit(Text1, .List1)
End With
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
With Form1
 .List1.ListIndex = .WordFilter1.Findit(Text1, .List1)
End With
KeyAscii = 0
End If
End Sub
