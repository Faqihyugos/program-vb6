VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFF00&
   Caption         =   "SEGITIGA"
   ClientHeight    =   5085
   ClientLeft      =   6675
   ClientTop       =   3735
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   4995
   Begin VB.CommandButton Command2 
      Caption         =   "SEBELUMNYA"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   360
      TabIndex        =   3
      Top             =   360
      Width           =   2895
   End
   Begin VB.CommandButton k 
      BackColor       =   &H000000FF&
      Caption         =   "KELUAR"
      Height          =   375
      Left            =   3600
      MaskColor       =   &H8000000F&
      TabIndex        =   2
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "TAMPILKAN"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim i, x, l As Integer
Dim j As String
List1.Clear
x = Val(Text2)
For i = 1 To x
f = ""
For q = 1 To i
f = f & q
Next
List1.AddItem (f)
Next i
End Sub

Private Sub Command2_Click()
form1.Show
Form2.Hide
End Sub

Private Sub k_Click()
End
End Sub
