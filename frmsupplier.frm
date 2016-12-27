VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmsupplier 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Input Data Supplier"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   150
      Left            =   4440
      Top             =   120
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   3720
      TabIndex        =   17
      Top             =   3120
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   1200
      TabIndex        =   15
      Top             =   3120
      Width           =   1215
   End
   Begin Project1.Button btnsimpan 
      Height          =   375
      Left            =   840
      TabIndex        =   8
      Top             =   2280
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
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   1110
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   680
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin Project1.Button btnbatal 
      Height          =   375
      Left            =   1920
      TabIndex        =   9
      Top             =   2280
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
      Left            =   3000
      TabIndex        =   10
      Top             =   2280
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
      Left            =   4080
      TabIndex        =   11
      Top             =   2280
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
   Begin MSDataGridLib.DataGrid Grid 
      Height          =   2295
      Left            =   120
      TabIndex        =   12
      Top             =   3720
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4048
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama"
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
      Left            =   2880
      TabIndex        =   16
      Top             =   3120
      Width           =   660
   End
   Begin VB.Label Label6 
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
      TabIndex        =   14
      Top             =   3120
      Width           =   585
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Height          =   225
      Left            =   2400
      TabIndex        =   13
      Top             =   2880
      Width           =   1200
   End
   Begin VB.Shape Shape2 
      Height          =   735
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   2880
      Width           =   5895
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C000&
      Height          =   615
      Left            =   720
      Shape           =   4  'Rounded Rectangle
      Top             =   2160
      Width           =   4455
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Telepon"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Top             =   1650
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Supplier"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1605
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Supplier"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1515
   End
End
Attribute VB_Name = "frmsupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rssup As New ADODB.Recordset
Dim strTemp, LenTemp, n, enTemp

Sub tampil()
Dim p
status
p = "select * from Supplier"
rssup.Open p, db
Set Grid.DataSource = rssup
Grid.Refresh
'DBCombo1.Clear
'Do While Not rs.EOF
 '  DBCombo1.AddItem rs!KodeBrg
  ' rs.MoveNext
'Loop
End Sub

Sub status()
Set rssup = New ADODB.Recordset
rssup.CursorLocation = adUseClient
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
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Then
    MsgBox "Isikan data secara lengkap !", vbCritical + vbOKOnly, "Peringatan !"
    Text1.SetFocus
    Else
Dim p
Select Case MsgBox("Apakah data sudah benar ?", vbYesNo, "Perhatian")
Case vbYes
    status
    p = "update Supplier set Kd_sup='" & Trim(Text1.Text) & "',Nama_sup='" & Trim(Text2.Text) & "',Alamat='" & Trim(Text3.Text) & "',Telepon='" & Trim(Text4.Text) & "' where Kd_sup='" & Trim(Grid.Columns(0)) & "'"
        
    rssup.Open p, db
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
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Then
    MsgBox "Pilih Data dahulu sebelum menghapus !", vbCritical + vbOKOnly, "Peringatan !"
    Else
Dim p
Select Case MsgBox("Apakah data akan dihapus ?", vbYesNo, "Perhatian")
Case vbYes
    status
    p = "delete * from Supplier where Kd_sup='" & Trim(Text1.Text) & "'"
    rssup.Open p, db
    tampil
    atur
    kosong
    Matikan
Case vbNo
End Select
End If
End Sub

Private Sub btnSimpan_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Then
    MsgBox "Isikan data secara lengkap !", vbCritical + vbOKOnly, "Peringatan !"
    Text1.SetFocus
    Else
    Dim X
    status
    X = "insert into Supplier(Kd_sup,Nama_sup,Alamat,Telepon) values ('" & Trim(Text1.Text) & "','" & Trim(Text2.Text) & "','" & Trim(Text3.Text) & "','" & Trim(Text4.Text) & "')"
    rssup.Open X, db
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
menuutama.StatusBar1.Panels(1).Text = "Input Data Supplier"
End Sub
Sub atur()
Grid.Columns(0).Caption = "Kode Supplier"
Grid.Columns(0).Width = "1200"
Grid.Columns(1).Caption = "Nama Supplier"
Grid.Columns(1).Alignment = dbgGeneral
Grid.Columns(1).Width = "2000"
Grid.Columns(2).Caption = "Alamat"
Grid.Columns(2).Width = "3000"
Grid.Columns(3).Caption = "Telepon"
Grid.Columns(3).Width = "1500"
End Sub
Sub kosong()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
menuutama.Enabled = True
menuutama.StatusBar1.Panels(1).Text = "Halaman Utama System"
End Sub

Private Sub Grid_DblClick()
Text1.Text = Trim(rssup("Kd_sup"))
Text2.Text = Trim(rssup("Nama_sup"))
Text3.Text = Trim(rssup("Alamat"))
Text4.Text = Trim(rssup("Telepon"))
bukakunci
btnedit.Enabled = True
btnhapus.Enabled = True
End Sub

Private Sub Text1_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then
    If Text1.Text = "" Then
        MsgBox "Isikan Kode Supplier dahulu !", vbCritical + vbOKCancel, "Peringatan"
        Else
    Dim p
    status
    p = "select * from Supplier where Kd_sup='" & Trim(Text1.Text) & "'"
    rssup.Open p, db
    If rssup.BOF Or rssup.EOF Then
        Text2.Enabled = True
        Text2.SetFocus
        Grid.Enabled = False
        Else
        MsgBox "Kode Supplier sudah ada !", vbCritical + vbOKOnly, "Peringatan !"
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
End Sub
Private Sub Text2_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then
    If Text2.Text = "" Then
        MsgBox "Isikan Nama Supplier dahulu !", vbCritical + vbOKOnly, "Peringatan !"
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
End Sub
Private Sub Text3_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then
    If Text3.Text = "" Then
        MsgBox "Isikan Alamat Supplier dahulu !", vbCritical + vbOKOnly, "Peringatan !"
        Text3.SetFocus
        Else
        Text4.Enabled = True
        Text4.SetFocus
    End If
End If
End Sub
Private Sub Text4_KeyPress(Keyascii As Integer)
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
Sub Cari(Kd_sup, Nama_sup)
Dim p
status
p = "select * from supplier where " & Kd_sup & " like '%" & Nama_sup & "%' order by Kd_sup"
rssup.Open p, db
Set Grid.DataSource = rssup
Grid.Refresh
End Sub

Private Sub Text5_Change()
If Trim(Text5) <> Empty Then
    Cari "Kd_sup", Text5.Text
Else
    tampil
End If
End Sub

Private Sub Text5_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
End Sub

Private Sub Text6_Change()
If Trim(Text6) <> Empty Then
    Cari "Nama_sup", Text6.Text
Else
    tampil
End If
End Sub

Private Sub Text6_KeyPress(Keyascii As Integer)
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
