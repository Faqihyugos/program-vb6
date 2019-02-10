VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H000080FF&
   Caption         =   "KASIR"
   ClientHeight    =   6585
   ClientLeft      =   5850
   ClientTop       =   2715
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   7875
   Begin VB.CommandButton Command3 
      Caption         =   "Keluar"
      Height          =   375
      Left            =   4800
      TabIndex        =   18
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Reset"
      Height          =   375
      Left            =   6360
      TabIndex        =   17
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cetak struk"
      Height          =   375
      Left            =   4800
      TabIndex        =   16
      Top             =   4680
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   405
      Left            =   2640
      TabIndex        =   14
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   405
      Left            =   2640
      TabIndex        =   13
      Top             =   3960
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   3180
      ItemData        =   "test.frx":0000
      Left            =   4800
      List            =   "test.frx":0002
      TabIndex        =   10
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   405
      Left            =   2640
      TabIndex        =   9
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   2640
      TabIndex        =   8
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   405
      Left            =   2640
      TabIndex        =   7
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   405
      Left            =   2640
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2640
      TabIndex        =   2
      Text            =   "Pilih"
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "KASIR TYAS SELULER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   19
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Kembalian"
      Height          =   375
      Left            =   840
      TabIndex        =   15
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Uang bayar"
      Height          =   375
      Left            =   840
      TabIndex        =   12
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label struck 
      BackStyle       =   0  'Transparent
      Caption         =   "Struk"
      Height          =   375
      Left            =   4920
      TabIndex        =   11
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Total harga"
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah barang"
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Harga barang"
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama barang"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Kode barang"
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
If Combo1 = "OPPA37" Then
Text1 = "OPPO A37"
Text2 = "1620000"
ElseIf Combo1 = "OPPF5" Then
Text1 = "OPPO F5"
Text2 = "3500000"
ElseIf Combo1 = "XIAR4X" Then
Text1 = "XIAOMI R4X"
Text2 = "1500000"
ElseIf Combo1 = "VIV9" Then
Text1 = "VIVO V9"
Text2 = "3500000"
End If
End Sub

Private Sub Command1_Click()
List1.AddItem ("Nama barang   : " + Text1)
List1.AddItem ("Harga barang  : Rp." + Text2)
List1.AddItem ("Jumlah barang : " + Text3)
List1.AddItem ("Total harga   : Rp." + Text4)
List1.AddItem ("Uang bayar    : Rp." + Text5)
List1.AddItem ("Kembalian     : Rp." + Text6)


End Sub

Private Sub Command2_Click()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
List1.Clear
Combo1.Text = ""

End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Form_Load()
Combo1.AddItem ("OPPA37")
Combo1.AddItem ("OPPF5")
Combo1.AddItem ("XIAR4X")
Combo1.AddItem ("VIV9")
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text4 = Val(Text3 * Text2)
End If

End Sub


Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text6 = Val(Text5 - Text4)
End If

End Sub
