VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form2 
   Caption         =   "Registrasi"
   ClientHeight    =   4995
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7065
   LinkTopic       =   "Form2"
   ScaleHeight     =   4995
   ScaleWidth      =   7065
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "Kembali ke Menu"
      Height          =   495
      Left            =   4080
      TabIndex        =   9
      Top             =   2160
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   720
      Top             =   3120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\semester 2\vb\10417082_P6_VB\database.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\semester 2\vb\10417082_P6_VB\database.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "login"
      Caption         =   "Data"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form2.frx":0000
      Height          =   1095
      Left            =   720
      TabIndex        =   8
      Top             =   3480
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   1931
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cari"
      Height          =   495
      Left            =   4080
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Tambah"
      Height          =   495
      Left            =   4080
      TabIndex        =   6
      Top             =   600
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2280
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   3720
      X2              =   3720
      Y1              =   480
      Y2              =   2880
   End
   Begin VB.Label Label3 
      Caption         =   "Status"
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Username"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub command1_click()
Adodc1.Recordset.Find "Nama_pengguna='" + Text1.Text + "'", , adSearchForward, 1
If Adodc1.Recordset.EOF Then
 MsgBox "Simpan Data"
Adodc1.Recordset.AddNew
Adodc1.Recordset!Nama_pengguna = Text1.Text
Adodc1.Recordset!Kata_sandi = Text2.Text
Adodc1.Recordset!Status = Combo1
Else
    MsgBox "Username sudah terdaftar"
End If
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.Find "Nama_pengguna='" + Text1.Text + "'", , adSearchForward, 1
If Adodc1.Recordset.EOF Then
    MsgBox "Gagal ditemukan"
Else
    MsgBox "Berhasil ditemukan"
End If
End Sub

Private Sub Command3_Click()
Form2.Hide
Form1.Show
End Sub

Private Sub DataGrid1_Click()
Text1.Text = Adodc1.Recordset!Nama_pengguna
Text2.Text = Adodc1.Recordset!Kata_sandi
End Sub

Private Sub Form_Load()
Combo1.AddItem ("Admin")
Combo1.AddItem ("kasir")
End Sub


