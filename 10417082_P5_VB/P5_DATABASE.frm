VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFF00&
   Caption         =   "Form1"
   ClientHeight    =   6915
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10830
   LinkTopic       =   "Form1"
   ScaleHeight     =   6915
   ScaleWidth      =   10830
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option2 
      Caption         =   "PEREMPUAN"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   19
      Top             =   2040
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "LAKI-LAKI"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   18
      Top             =   2040
      Width           =   1455
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8040
      TabIndex        =   17
      Text            =   "TAHUN"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6600
      TabIndex        =   16
      Text            =   "BULAN"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5160
      TabIndex        =   15
      Text            =   "HARI"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   5160
      TabIndex        =   11
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   5160
      TabIndex        =   10
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   1440
      TabIndex        =   9
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   1440
      TabIndex        =   5
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "HAPUS"
      Height          =   495
      Left            =   9120
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SIMPAN"
      Height          =   495
      Left            =   7680
      TabIndex        =   1
      Top             =   3120
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   240
      Top             =   6240
      Width           =   10200
      _ExtentX        =   17992
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\semester 2\vb\10417082_P5_VB\DATABASE.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\semester 2\vb\10417082_P5_VB\DATABASE.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Table1"
      Caption         =   "DATA YANG TERSIMPAN"
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
      Bindings        =   "P5_DATABASE.frx":0000
      Height          =   2295
      Left            =   240
      TabIndex        =   0
      Top             =   3960
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   4048
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
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DATA MAHASISWA STMIK JAKARTA 2018"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   240
      TabIndex        =   22
      Top             =   480
      Width           =   9015
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "JENIS KELAMIN"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   21
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "TANGGAL LAHIR"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   20
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TAHUN MASUK"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   14
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "KELAS"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   13
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "JENJANG"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "JURUSAN"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NAMA"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NPM"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   1440
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Text2.Text = "" Or Text5.Text = "" Or Text1.Text = "" Or Combo1.Text = "" Or Combo1.Text = "HARI" Or Combo2.Text = "" Or Combo2.Text = "BULAN" Or Combo3.Text = "" Or Combo3.Text = "TAHUN" Then
MsgBox "TERDAPAT DATA KOSONG", vbCritical, "PERINGATAN"
Else
Adodc1.Recordset.AddNew
Adodc1.Recordset!NPM = Text1.Text
Adodc1.Recordset!NAMA = Text2.Text
Adodc1.Recordset!JURUSAN = Text3.Text
Adodc1.Recordset!JENJANG = Text4.Text
Adodc1.Recordset!KELAS = Text5.Text
Adodc1.Recordset!TAHUN_MASUK = Text6.Text
Adodc1.Recordset!TANGGAL_LAHIR = Combo1 + Combo2 + Combo3
If Option1.Value = True Then
Adodc1.Recordset!JENIS_KELAMIN = Option1.Caption
Else
Adodc1.Recordset!JENIS_KELAMIN = Option2.Caption
End If
Adodc1.Recordset.Update
MsgBox "DATA BERHASIL DISIMPAN", vbInformation, "BERHASIL"
End If
End Sub

Private Sub Command2_Click()
If MsgBox("Anda Yakin Ingin Menghapus", vbYesNo, "DATA") = vbYes Then
If Adodc1.Recordset.RecordCount <> 0 Then Adodc1.Recordset.Delete
MsgBox "Data Berhasil di Hapus", vbInformation, "DATA"
ElseIf vbNo Then
MsgBox "DIBATALKAN !", vbCritical, "Keluar"
End If
End Sub

Private Sub Form_Load()
For a = 1 To 31
Combo1.AddItem a
Next a

For B = 1 To 12
Combo2.AddItem (MonthName(B))
Next B

For c = 1989 To Year(Now)
Combo3.AddItem (Str(c))
Next c

End Sub


Private Sub Text1_Change()
If Len(Text1.Text) > 8 Then
MsgBox "LEBIH DARI 8 INPUTAN", vbCritical, "SALAH"
Text1.Text = Left(Text1.Text, 8)
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

a = Left(Text1.Text, 3)
If a = 104 Then Text3 = ("SISTEM INFORMASI")
If a = 204 Then Text3 = ("SISTEM KOMPUTER")
If a = 304 Then Text3 = ("MANAJEMEN INFORMASI")
If a = 404 Then Text3 = ("TEKNIK KOMPUTER")

c = Left(Text1.Text, 1)
If c = 1 Then Text4 = ("S1")
If c = 2 Then Text4 = ("S1")
If c = 3 Then Text4 = ("D3")
If c = 4 Then Text4 = ("D3")

d = Mid(Text1.Text, 4, 2)
If d >= 78 Then Text6 = ("19" & d)
If d <= 78 Then Text6 = ("20" & d)

Text2.SetFocus
End If
End Sub
