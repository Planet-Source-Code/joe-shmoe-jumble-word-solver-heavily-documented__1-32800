VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Un-Jumbler"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   Icon            =   "frmJumble2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   5280
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Possible Matches"
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   5055
      Begin VB.ListBox List1 
         Height          =   1035
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2640
      MaxLength       =   9
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   5160
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   5160
      X2              =   5160
      Y1              =   1440
      Y2              =   120
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   5160
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   1440
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      Height          =   375
      Left            =   240
      Top             =   840
      Width           =   4695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "Find Possible Words"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   15.75
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   4695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Jumbled Word"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2235
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label2_Click()
If Len(Text1.Text) < 2 Or Len(Text1.Text) > 9 Then 'if the string length is out of range then do nothing
    MsgBox "Word length must be between 2 and 9 characters", vbExclamation, "UnJumbler"
    Exit Sub
End If
frmSearching.Show 'show the searching form
DoEvents 'give the searching form time to load
List1.Clear 'clear the list
List1.Visible = False 'make the list invisible
Dim CheckWord As String 'string the user entered
Dim ListWord As String 'word from the file
Dim LenTimes As Long 'used to compare the lengths of the 2 strings
CheckWord = LCase(Text1.Text)
Open App.Path & "\acd9.txt" For Input As #1 'open the word list
Do While Not EOF(1)
ListWord = "" 'its the beginning of the loop, so ListWord, and LenTimes are reset
LenTimes = 0
Line Input #1, ListWord 'get the word from the file
LCase (ListWord) 'for easier comparison
If Len(ListWord) = Len(CheckWord) Then 'we only need to check if the words are the same lenght
    For i = 97 To 122 '97-122 are the lowercase Asc() for a-z
        c = FindOccur(CheckWord, Chr(i)) 'find the number of times Chr(i) is in the string CheckWord
        l = FindOccur(ListWord, Chr(i)) 'same thing as the previous statement
        If c <> l Then 'if the letter occurs more times in one string than the other then it isn't possible for it to be a match
            Exit For 'exit the For loop and start over
        End If
        If c <> 0 And l <> 0 And c = l Then 'if the 2 words have atleast one character in common then add the number of times the letter [Chr(i)] occured
            LenTimes = LenTimes + c
        End If
    If LenTimes = Len(CheckWord) Then 'if they have the same number of matching letters then its a match
        LenTimes = 0
        lw = UCase(Left(ListWord, 1)) 'upper case the first letter
        rw = LCase(Right(ListWord, Len(ListWord) - 1)) 'lowercase the rest of the word
        List1.AddItem lw + rw 'and the string to the list box
    End If
    Next i
End If
Loop
Close 'close the file
List1.Visible = True 'show the list
Unload frmSearching
End Sub

Public Function FindOccur(Word As String, Character As String) As Integer
ReDim letters(Len(Word))
For i = 0 To Len(Word) - 1
    letters(i) = Mid(Word, i + 1, 1) 'stips the WORD apart to an array
    If Character = letters(i) Then 'if the CHARACTER is in the word then add it to the total
        Times = Times + 1
    End If
Next i
FindOccur = Times 'return the value
End Function
