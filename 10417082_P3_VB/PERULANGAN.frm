VERSION 5.00
Begin VB.Form form1 
   BackColor       =   &H00FFFF00&
   Caption         =   "PROGRAM PERULANGAN"
   ClientHeight    =   5100
   ClientLeft      =   7080
   ClientTop       =   3540
   ClientWidth     =   4380
   LinkTopic       =   "Form1"
   ScaleHeight     =   5100
   ScaleWidth      =   4380
   Begin VB.CommandButton Command5 
      BackColor       =   &H00000000&
      Caption         =   "SELANJUTNYA"
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H000000FF&
      Caption         =   "KELUAR"
      Height          =   375
      Left            =   840
      TabIndex        =   7
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   2280
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   840
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "DO UNTIL"
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "DO WHILE"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "FOR NEXT"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   840
      TabIndex        =   0
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PROGRAM PERULANGAN"
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   360
      Width           =   3015
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A, B As Integer

Private Sub Command1_Click()
List1.Clear
A = Text1
B = Text2
For A = A To B
List1.AddItem (A)
Next A
End Sub

Private Sub Command2_Click()
List1.Clear
A = Text1
B = Text2
Do While A <= B
List1.AddItem (A)
A = A + 1
Loop
End Sub

Private Sub Command3_Click()
List1.Clear
A = Text1
B = Text2
Do Until A > B
List1.AddItem (A)
A = A + 1
Loop
End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub Command5_Click()
Form2.Show
form1.Hide
End Sub
