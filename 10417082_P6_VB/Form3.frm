VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Menu Login"
   ClientHeight    =   3465
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6090
   LinkTopic       =   "Form3"
   ScaleHeight     =   3465
   ScaleWidth      =   6090
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check1 
      Caption         =   "Show"
      Height          =   495
      Left            =   4200
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "password"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Username"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A As Byte
Dim Koneksi As New ADODB.Connection
Dim RSUser As ADODB.Recordset
Sub BukaDB()
 Set Koneksi = New ADODB.Connection
 Set RSUser = New ADODB.Recordset
 Koneksi.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\database.mdb"
End Sub
Private Sub Check1_Click()
If Check1.Value = 1 Then
Text2.PasswordChar = ""
Else
Text2.PasswordChar = "*"
End If
End Sub
Private Sub command1_click()
Call BukaDB
        RSUser.Open "Select * from login where Nama_pengguna ='" & Text1 & "' and Kata_sandi='" & Text2 & "'", Koneksi
        If RSUser.EOF Then
        A = A + 1
            If 1 - A = 0 Then
                MsgBox "Kesempatan ke " & A & " Salah" & Chr(13) & _
                        "User dan Password tidak dikenal"
                Text1.SetFocus
            ElseIf 2 - A = 0 Then
                MsgBox "Kesempatan ke " & A & " Salah" & Chr(13) & _
                        "User dan Password tidak dikenal"
                Text1.SetFocus
            ElseIf 3 - A = 0 Then
                MsgBox "Kesempatan ke " & A & " Salah" & Chr(13) & _
                        "User dan Password tidak dikenal" & Chr(13) & _
                        "Kesempatan habis, Ulangi dari awal"
                Form3.Hide
                Form1.Show
            End If
        Else
            If RSUser!Status = "admin" Then
            Form3.Hide
            Form4.Show
            MsgBox "Berhasil login sebagai admin"
            ElseIf RSUser!Status = "kasir" Then
            Form3.Hide
            Form5.Show
            MsgBox "Berhasil login sebagai kasir"
            End If
        End If
End Sub
