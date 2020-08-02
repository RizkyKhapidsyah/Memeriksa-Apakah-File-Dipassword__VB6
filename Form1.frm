VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Memeriksa Apakah File Dipassword"
   ClientHeight    =   1695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7785
   LinkTopic       =   "Form1"
   ScaleHeight     =   1695
   ScaleWidth      =   7785
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Browse.."
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Periksa"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function Password_Check(Path As String) As _
String
Dim db As DAO.Database
  If Dir(Path) = "" Then
  'Kembalikan 0 jika file tidak ada
     Password_Check = "0"
     Exit Function
  End If

  If Right(Path, 3) = "mdb" Then
      On Error GoTo errorline
      Set db = OpenDatabase(Path)
      Password_Check = "False"
      MsgBox "File " & Path & "" & Chr(13) & _
             "adalah file yang tidak dipassword!", _
             vbInformation, "Akses Diterima"
      db.Close
      Exit Function
  ElseIf Right(Path, 3) = "xls" Then
      On Error GoTo errorline
      Set db = OpenDatabase(Path, True, _
               False, "Excel 5.0")
      Password_Check = "False"
      MsgBox "File " & Path & "" & Chr(13) & _
             "adalah file yang tidak dipassword!", _
             vbInformation, "Akses Diterima"
      db.Close
      Exit Function
  Else
      'Asumsikan bukan file yang valid jika ekstensinya
      'bukan xls atau mdb seperti di atas
      Password_Check = "0"
      MsgBox "File " & Path & "" & Chr(13) & _
             "adalah file yang tidak dipassword!", _
              vbInformation, "Akses Diterima"
     Exit Function
  End If
errorline:
    Password_Check = "True"
    MsgBox "File " & Path & "" & Chr(13) & _
           "adalah file yang dipassword!", _
            vbCritical, "Akses Ditolak"
    Exit Function
End Function

Private Sub Command1_Click()  'Untuk memeriksa apakah
                              'file dipassword?
  If CommonDialog1.FileName = "" Then
     MsgBox "Pilih nama file dari tombol Browse...!", _
             vbCritical, "Pilih Nama File"
     Exit Sub
  Else
     Password_Check (CommonDialog1.FileName)
  End If
End Sub

Private Sub Command2_Click()   'Untuk memilih file yang
                               'akan diperiksa
On Error Resume Next
  With CommonDialog1
   .Filter = "Semua Files|*.*"
   .DialogTitle = "Ambil Nama File..."
   .ShowOpen
  End With
  Label1.Caption = CommonDialog1.FileName
End Sub


