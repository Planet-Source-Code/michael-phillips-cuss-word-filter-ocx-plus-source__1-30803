VERSION 5.00
Begin VB.UserControl WordFilter 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "WordFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private FWord() As String
Public WordCount As Integer

Public Function AddWord(TheNewWord As String) As Boolean
Dim U As Integer
If Trim(TheNewWord) = "" Then
   AddWord = False
   Exit Function
End If
For U = 1 To WordCount
    If LCase(TheNewWord) = LCase(FWord(U)) Then
       MsgBox TheNewWord & " That word is already in the list!", vbInformation, "Word Filter"
       AddWord = False
       GoTo KK:
    End If
Next
   ReDim Preserve FWord(WordCount + 1)
   WordCount = WordCount + 1
   FWord(WordCount) = TheNewWord
   AddWord = True
KK:
End Function

Function Findit(xOx As String, ByRef List As Object) As Integer
On Error GoTo fff:
Dim E As Integer
For E = 0 To List.ListCount
    If LCase(xOx) = LCase(List.List(E)) Then
       Findit = E
       List.ListIndex = Findit
       Exit For
    End If
Next
fff:
End Function

Sub RemoveWord(TheWord As String, _
   Optional FromList As Boolean, _
 Optional List As Object)
''''''''
On Error Resume Next
Dim U As Integer
For U = 1 To WordCount
    If LCase(TheWord) = LCase(FWord(U)) Then
       FWord(U) = ""
       WordCount = Val(WordCount) - 1
       ReDim Preserve FWord(WordCount)
       GoTo KK:
    End If
Next
KK:
 If FromList Then
    Findit TheWord, List
    DoEvents
    List.RemoveItem List.ListIndex
 End If

    
End Sub

Function CheckWords(TheLine As String, _
Optional ReplaceWord As Boolean, Optional NEWLINE As String, _
Optional MyWay As Boolean, Optional Thereway As Boolean, _
Optional ThereChr As String) As Boolean

Dim i As Integer
Dim TheTemp As String
TheTemp = LCase(TheLine)
'''''
For i = 1 To WordCount
    If InStr(TheTemp, LCase(FWord(i))) > 0 Then
       CheckWords = True
       If ReplaceWord Then
       '
       If MyWay Then
         Dim TIT As String
         TIT = RCWord(FWord(i))
         TheTemp = Replace(LCase(TheTemp), LCase(FWord(i)), TIT)
         GoTo KK:
       End If
       '
       If Thereway Then
         TheTemp = Replace(LCase(TheTemp), LCase(FWord(i)), String(Len(FWord(i)), ThereChr))
         GoTo KK:
       End If
       '
         TheTemp = Replace(LCase(TheTemp), LCase(FWord(i)), String(Len(FWord(i)), "#"))
       End If
       '
    End If
KK:
Next
'''''
If CheckWords = True Then
    NEWLINE = TheTemp
    Exit Function
End If
'''''
CheckWords = False
End Function


Function RCWord(BadWord As String) As String
Dim aTemp As String
aTemp = LCase(BadWord)
If aTemp = "fucked" Then RCWord = "flocked": Exit Function
If aTemp = "fucker" Then RCWord = "flocker": Exit Function
If aTemp = "fuck" Then RCWord = "flock": Exit Function
If aTemp = "bitch" Then RCWord = "female dog": Exit Function
If aTemp = "bastard" Then RCWord = "jubba jubba": Exit Function
If aTemp = "asshole" Then RCWord = "poo shoot": Exit Function
If aTemp = "ass" Then RCWord = "booty": Exit Function
If aTemp = "cunt" Then RCWord = "endless hole": Exit Function
If aTemp = "nigger" Then RCWord = "dubba dubba": Exit Function
If aTemp = "cocksucker" Then RCWord = "water hazard": Exit Function
If aTemp = "motherfucker" Then RCWord = "mommy flocker": Exit Function
If aTemp = "slut" Then RCWord = "rut": Exit Function
If aTemp = "whore" Then RCWord = "female worker": Exit Function
If aTemp = "dick" Then RCWord = "pee wee": Exit Function
If aTemp = "cock" Then RCWord = "big guy": Exit Function
If aTemp = "pussy" Then RCWord = "kitty": Exit Function
If aTemp = "cum" Then RCWord = "splatter": Exit Function
If aTemp = "rape" Then RCWord = "force": Exit Function
If aTemp = "faggot" Then RCWord = "Maggot": Exit Function
If aTemp = "fag" Then RCWord = "rag": Exit Function
If aTemp = "anal" Then RCWord = "intruder": Exit Function
If aTemp = "tit" Then RCWord = "breast": Exit Function
If aTemp = "porno" Then RCWord = "adult movie": Exit Function
If aTemp = "p0rn" Then RCWord = "fl0wer": Exit Function
If aTemp = "porn" Then RCWord = "flower": Exit Function
If aTemp = "suck" Then RCWord = "flower": Exit Function
End Function

Public Sub SaveList(ByRef TheList As Object, FileName As String)
    On Error Resume Next
    Dim Save As Long
    Dim fFile As Integer
    fFile = FreeFile
    Open FileName For Output As fFile
    For Save = 0 To TheList.ListCount - 1
        Print #fFile, TheList.List(Save)
        DoEvents
    Next Save
    Close fFile
End Sub

Public Sub ListLoad(FileName As String, Optional ByRef TheList As Object)
    On Error Resume Next
    Dim TheContents As String
    Dim fFile As Integer
    fFile = FreeFile
    Open FileName For Input As fFile
    Do
        Line Input #fFile, TheContents$
        If TheContents$ = "" Then GoTo KK:
        AddWord TheContents$
        TheList.AddItem TheContents$
KK:
    Loop Until EOF(fFile)
    Close fFile
End Sub

Private Sub UserControl_Initialize()
'MsgBox "Coded by Michael Phillips Please take 1 minute and VOTE!", vbInformation, "Word Filter OCX V1.0"
End Sub
