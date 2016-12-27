VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmpegawai 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Input Data Pegawai"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   150
      Left            =   4440
      Top             =   120
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   3960
      TabIndex        =   19
      Top             =   3360
      Width           =   2055
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   1200
      TabIndex        =   18
      Top             =   3360
      Width           =   1215
   End
   Begin Project1.Button btnsimpan 
      Height          =   375
      Left            =   960
      TabIndex        =   11
      Top             =   2520
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Simpan"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin MSDataGridLib.DataGrid Grid 
      Height          =   2415
      Left            =   120
      TabIndex        =   10
      Top             =   3960
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   4260
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
            LCID            =   1057
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
            LCID            =   1057
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
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   1500
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   660
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   240
      Width           =   975
   End
   Begin Project1.Button btnbatal 
      Height          =   375
      Left            =   2040
      TabIndex        =   12
      Top             =   2520
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Batal"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin Project1.Button btnedit 
      Height          =   375
      Left            =   3120
      TabIndex        =   13
      Top             =   2520
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Edit"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin Project1.Button btnhapus 
      Height          =   375
      Left            =   4200
      TabIndex        =   14
      Top             =   2520
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Hapus"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama "
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3120
      TabIndex        =   17
      Top             =   3360
      Width           =   720
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kode"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   480
      TabIndex        =   16
      Top             =   3360
      Width           =   585
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Pencarian Data"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   15
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      Height          =   735
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   6255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   615
      Left            =   840
      Shape           =   4  'Rounded Rectangle
      Top             =   2400
      Width           =   4455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Telepon"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1900
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1500
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Jabatan"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Pegawai"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   680
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Pegawai"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
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
      Width           =   1815
   End
End
Attribute VB_Name = "frmpegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rspeg As New ADODB.Recordset
Dim strTemp, LenTemp, n, enTemp

Sub tampil()
Dim p
status
p = "select * from Pegawai"
rspeg.Open p, db
Set Grid.DataSource = rspeg
Grid.Refresh
'DBCombo1.Clear
'Do While Not rs.EOF
 '  DBCombo1.AddItem rs!KodeBrg
  ' rs.MoveNext
'Loop
End Sub

Sub status()
Set rspeg = New ADODB.Recordset
rspeg.CursorLocation = adUseClient
End Sub

Private Sub btnbatal_Click()
tampil
atur
kosong
Matikan
Text1.SetFocus
btnsimpan.Enabled = False
btnedit.Enabled = False
btnhapus.Enabled = False
Grid.Enabled = True
End Sub

Private Sub btnedit_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Then
    MsgBox "Isikan data secara lengkap !", vbCritical + vbOKOnly, "Peringatan !"
    Text1.SetFocus
    Else
Dim p
Select Case MsgBox("Apakah data sudah benar ?", vbYesNo, "Perhatian")
Case vbYes
    status
    p = "update Pegawai set Kd_peg='" & Trim(Text1.Text) & "',Nama_peg='" & Trim(Text2.Text) & "',Jabatan='" & Trim(Text3.Text) & "',Alamat='" & Trim(Text4.Text) & "',Telepon='" & Trim(Text5.Text) & "' where Kd_peg='" & Trim(Grid.Columns(0)) & "'"
        
    rspeg.Open p, db
    MsgBox "Data telah berubah dan tersimpan.", vbOKOnly, "Informasi"
    tampil
    atur
    kosong
    Matikan
    Text1.SetFocus
    btnedit.Enabled = False
    btnhapus.Enabled = True
Case vbNo
End Select
End If
End Sub

Private Sub btnhapus_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Then
    MsgBox "Pilih Data dahulu sebelum menghapus !", vbCritical + vbOKOnly, "Peringatan !"
    Else
Dim p
Select Case MsgBox("Apakah data akan dihapus ?", vbYesNo, "Perhatian")
Case vbYes
    status
    p = "delete * from Pegawai where Kd_peg='" & Trim(Text1.Text) & "'"
    rspeg.Open p, db
    tampil
    atur
    kosong
    Matikan
Case vbNo
End Select
End If
End Sub

Private Sub btnSimpan_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Then
    MsgBox "Isikan data secara lengkap !", vbCritical + vbOKOnly, "Peringatan !"
    Text1.SetFocus
    Else
    Dim X
    status
    X = "insert into Pegawai(Kd_peg,Nama_peg,Jabatan,Alamat,Telepon) values ('" & Trim(Text1.Text) & "','" & Trim(Text2.Text) & "','" & Trim(Text3.Text) & "','" & Trim(Text4.Text) & "','" & Trim(Text5.Text) & "')"
    rspeg.Open X, db
    MsgBox "Data sudah tersimpan !", vbInformation, "Informasi"
    kosong
    tampil
    atur
    Matikan
    btnsimpan.Enabled = False
    Grid.Enabled = True
    Text1.SetFocus
End If
End Sub



Private Sub Form_Load()
strTemp = Me.Caption
    n = 1
Call SetFormCenter(Me)
koneksi
tampil
'tampilcombo
atur
Matikan
btnsimpan.Enabled = False
btnedit.Enabled = False
btnhapus.Enabled = False
menuutama.StatusBar1.Panels(1).Text = "Input Data Pegawai"
End Sub
Sub atur()
Grid.Columns(0).Caption = "Kode Pegawai"
Grid.Columns(0).Width = "1200"
Grid.Columns(1).Caption = "Nama Pegawai"
Grid.Columns(1).Alignment = dbgGeneral
Grid.Columns(1).Width = "2000"
Grid.Columns(2).Caption = "Jabatan"
Grid.Columns(2).Width = "1500"
Grid.Columns(3).Caption = "Alamat"
Grid.Columns(3).Width = "3000"
Grid.Columns(4).Caption = "Telepon"
Grid.Columns(4).Width = "1500"
End Sub
Sub kosong()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
menuutama.Enabled = True
menuutama.StatusBar1.Panels(1).Text = "Halaman Utama System"
End Sub

Private Sub Grid_DblClick()
Text1.Text = Trim(rspeg("Kd_peg"))
Text2.Text = Trim(rspeg("Nama_peg"))
Text3.Text = Trim(rspeg("Jabatan"))
Text4.Text = Trim(rspeg("Alamat"))
Text5.Text = Trim(rspeg("Telepon"))
bukakunci
btnedit.Enabled = True
btnhapus.Enabled = True
End Sub

Private Sub Text1_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then
    If Text1.Text = "" Then
        MsgBox "Isikan Kode Pegawai dahulu !", vbCritical + vbOKCancel, "Peringatan"
        Else
    Dim p
    status
    p = "select * from Pegawai where Kd_peg='" & Trim(Text1.Text) & "'"
    rspeg.Open p, db
    If rspeg.BOF Or rspeg.EOF Then
        Text2.Enabled = True
        Text2.SetFocus
        Grid.Enabled = False
        Else
        MsgBox "Kode Pegawai sudah ada !", vbCritical + vbOKOnly, "Peringatan !"
        Matikan
        Text1.Text = ""
        Text1.SetFocus
    End If
End If
End If
End Sub
Sub Matikan()
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
End Sub
Private Sub Text2_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then
    If Text2.Text = "" Then
        MsgBox "Isikan Nama Pegawai dahulu !", vbCritical + vbOKOnly, "Peringatan !"
        Text2.SetFocus
        Else
        Text3.Enabled = True
        Text3.SetFocus
    End If
End If
End Sub

Sub bukakunci()
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
End Sub
Private Sub Text3_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then
    If Text3.Text = "" Then
        MsgBox "Isikan Jabatan Pegawai dahulu !", vbCritical + vbOKOnly, "Peringatan !"
        Text3.SetFocus
        Else
        Text4.Enabled = True
        Text4.SetFocus
    End If
End If
End Sub
Private Sub Text4_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then
    If Text4.Text = "" Then
        MsgBox "Isikan Jabatan Pegawai dahulu !", vbCritical + vbOKOnly, "Peringatan !"
        Text3.SetFocus
        Else
        Text5.Enabled = True
        Text5.SetFocus
    End If
End If
End Sub
Private Sub Text5_KeyPress(Keyascii As Integer)
If Not (Keyascii >= vbKey0 And Keyascii <= vbKey9 Or Keyascii = vbKeyBack) Then
If Keyascii = 13 Then
    btnsimpan.Enabled = True
    btnhapus.Enabled = False
Else
Dim a
Beep
a = MsgBox("Hanya Bisa Diisi Angka", vbInformation, "Peringatan")
Keyascii = 0
   End If
   End If
End Sub
Sub Cari(Kd_peg, Nama_peg)
Dim p
status
p = "select * from pegawai where " & Kd_peg & " like '%" & Nama_peg & "%' order by Kd_peg"
rspeg.Open p, db
Set Grid.DataSource = rspeg
Grid.Refresh
End Sub

Private Sub Text6_Change()
If Trim(Text6) <> Empty Then
    Cari "Kd_peg", Text6.Text
Else
    tampil
End If
End Sub

Private Sub Text6_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
End Sub

Private Sub Text7_Change()
If Trim(Text7) <> Empty Then
    Cari "Nama_peg", Text7.Text
Else
    tampil
End If
End Sub

Private Sub Text7_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
End Sub

Private Sub Timer1_Timer()
enTemp = Len(strTemp)
    Dim Form As String
    LenTemp = Len(strTemp)
    Me.Caption = Left(strTemp, n) + "_"
    n = n + 1
    If n > LenTemp Then
        n = 1
    End If
End Sub
