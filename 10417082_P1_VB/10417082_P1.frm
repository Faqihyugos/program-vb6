VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "kasir sederhana"
   ClientHeight    =   6330
   ClientLeft      =   4425
   ClientTop       =   2505
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   Picture         =   "10417082_P1.frx":0000
   ScaleHeight     =   11.165
   ScaleMode       =   7  'Centimeter
   ScaleWidth      =   16.828
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      Caption         =   "Menghitung"
      Height          =   495
      Left            =   1200
      TabIndex        =   11
      Top             =   4680
      Width           =   1695
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   375
      Left            =   7080
      TabIndex        =   10
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   375
      Left            =   7080
      TabIndex        =   9
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   7080
      TabIndex        =   8
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "Selesai pembayaran"
      Height          =   495
      Left            =   7080
      TabIndex        =   7
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Refresh"
      Height          =   495
      Left            =   5640
      TabIndex        =   6
      Top             =   4080
      Width           =   975
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Input"
      Height          =   3255
      Left            =   600
      TabIndex        =   12
      Top             =   1200
      Width           =   3135
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "Harga"
         Height          =   255
         Left            =   1920
         TabIndex        =   22
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "Harga"
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "Jumlah"
         Height          =   255
         Left            =   1920
         TabIndex        =   20
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "Jumlah"
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "Makanan"
         Height          =   255
         Left            =   1800
         TabIndex        =   18
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "Minuman"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "Uang bayar"
      Height          =   375
      Left            =   5640
      TabIndex        =   17
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "Kembalian"
      Height          =   375
      Left            =   5640
      TabIndex        =   16
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "Harga total"
      Height          =   375
      Left            =   5640
      TabIndex        =   15
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "Kasir penjualan warung nasi"
      BeginProperty Font 
         Name            =   "NeoTech"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   14
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.Text = (" ")
Text2.Text = (" ")
Text3.Text = (" ")
Text4.Text = (" ")
Text5.Text = (" ")
Text6.Text = (" ")
Text7.Text = (" ")
Text8.Text = (" ")
Text9.Text = (" ")
End Sub

Private Sub Command2_Click()
Text8.Text = Val(Text7 - Text9)
End Sub

Private Sub Command3_Click()
Text9.Text = Val(Text3 * Text2) + Val(Text6 * Text5)
End Sub

