VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5445
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   5745
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Keluar"
      Height          =   375
      Left            =   3480
      TabIndex        =   11
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   4080
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Proses"
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   3480
      Width           =   975
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Kanan"
      Height          =   255
      Left            =   3720
      TabIndex        =   7
      Top             =   2880
      Width           =   975
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Tengah"
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   2880
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Kiri"
      Height          =   315
      Left            =   1080
      TabIndex        =   5
      Top             =   2880
      Width           =   975
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   3720
      TabIndex        =   4
      Top             =   2280
      Width           =   975
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2400
      TabIndex        =   3
      Top             =   2280
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1080
      TabIndex        =   2
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pilihan"
      Height          =   2655
      Left            =   720
      TabIndex        =   12
      Top             =   1920
      Width           =   4215
      Begin VB.CommandButton Command3 
         Caption         =   "Reset"
         Height          =   375
         Left            =   1800
         TabIndex        =   14
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Hasil manipulasi"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   2280
         Width           =   1335
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "MANIPULASI STRING"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   10
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Masukkan kata"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   1440
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim j, k, l As String

Private Sub Command1_Click()
If Check1.Value = 1 Then
j = Left(Text1.Text, Combo1.Text)
f = " "
Else
j = ""
End If

If Check2.Value = 1 Then
k = Mid(Text1.Text, Combo1.Text, Combo2.Text)
q = " "
Else
k = ""
End If

If Check3.Value = 1 Then
l = Right(Text1.Text, Combo3.Text)
Else
l = ""
End If


Text2.Text = j & f & k & q & l
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
Combo1.Text = ""
Combo2.Text = ""
Combo3.Text = ""
Text1.Text = ""
Text2.Text = ""
Text1.SetFocus
End Sub

Private Sub Form_Load()
For a = 1 To 10
Combo1.AddItem a
Combo2.AddItem a
Combo3.AddItem a
Next a
End Sub

