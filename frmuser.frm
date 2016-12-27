VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmuser 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Input Data User"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1320
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1800
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Timer Timer1 
      Left            =   4080
      Top             =   120
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   4320
      TabIndex        =   21
      Top             =   3360
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1560
      TabIndex        =   19
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Simpan"
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
      Height          =   2055
      Left            =   120
      TabIndex        =   10
      Top             =   3960
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   3625
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
   Begin VB.ComboBox DBCombo1 
      Height          =   315
      Left            =   1920
      TabIndex        =   9
      Top             =   650
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   1870
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   1450
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   1050
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   240
      Width           =   1095
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
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
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
      Left            =   3000
      TabIndex        =   20
      Top             =   3360
      Width           =   1125
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kode User"
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
      Left            =   240
      TabIndex        =   18
      Top             =   3360
      Width           =   1155
   End
   Begin VB.Label Label8 
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
      Left            =   2520
      TabIndex        =   17
      Top             =   3120
      Width           =   1200
   End
   Begin VB.Shape Shape2 
      Height          =   735
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   6135
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   4560
      TabIndex        =   16
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Pegawai"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3120
      TabIndex        =   15
      Top             =   600
      Width           =   1380
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
      Caption         =   "Status"
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
      Top             =   1870
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
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
      Top             =   1050
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      Top             =   1470
      Width           =   1095
   End
   Begin VB.Label Label2 
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
      TabIndex        =   1
      Top             =   650
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Kode User"
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
      Width           =   1455
   End
End
Attribute VB_Name = "frmuser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim strTemp, LenTemp, n, enTemp
Dim Data As Database

Sub tampil()
Dim p
status
p = "select * from [user]"
RS.Open p, db
Set Grid.DataSource = RS
Grid.Refresh
'DBCombo1.Clear
'Do While Not rs.EOF
 '  DBCombo1.AddItem rs!KodeBrg
  ' rs.MoveNext
'Loop
End Sub

Sub status()
Set RS = New ADODB.Recordset
RS.CursorLocation = adUseClient
Set rs2 = New ADODB.Recordset
rs2.CursorLocation = adUseClient
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
Label7.Caption = ""
End Sub

Private Sub btnedit_Click()
If Text1.Text = "" Or DBCombo1.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Then
    MsgBox "Isikan data secara lengkap !", vbCritical + vbOKOnly, "Peringatan !"
    Text1.SetFocus
    Else
Dim p
Select Case MsgBox("Apakah data sudah benar ?", vbYesNo, "Perhatian")
Case vbYes
       Data1.Recordset.Edit
    Data1.Recordset!Kd_user = Text1.Text
    Data1.Recordset!Kd_peg = DBCombo1.Text
    Data1.Recordset!Username = Text3.Text
    Data1.Recordset!Password = Text4.Text
    Data1.Recordset!status = Text5.Text
    Data1.Recordset.Update
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
If Text1.Text = "" Or Text3.Text = "" Or DBCombo1.Text = "" Or Text4.Text = "" Or Text5.Text = "" Then
    MsgBox "Pilih Data dahulu sebelum menghapus !", vbCritical + vbOKOnly, "Peringatan !"
    Else
Dim p
Select Case MsgBox("Apakah data akan dihapus ?", vbYesNo, "Perhatian")
Case vbYes
    status
    p = "delete * from [user] where Kd_user='" & Trim(Text1.Text) & "'"
    RS.Open p, db
    tampil
    atur
    kosong
    Matikan
Case vbNo
End Select
End If
End Sub

Private Sub btnSimpan_Click()
If Text1.Text = "" Or Text3.Text = "" Or DBCombo1.Text = "" Or Text4.Text = "" Or Text5.Text = "" Then
    MsgBox "Isikan data secara lengkap !", vbCritical + vbOKOnly, "Peringatan !"
    Text1.SetFocus
    Else
    Dim X
    Data1.Recordset.AddNew
    Data1.Recordset!Kd_user = Text1.Text
    Data1.Recordset!Kd_peg = DBCombo1.Text
    Data1.Recordset!Username = Text3.Text
    Data1.Recordset!Password = Text4.Text
    Data1.Recordset!status = Text5.Text
    Data1.Recordset.Update
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




Private Sub DBCombo1_Click()
status
Dim p
p = "select * from pegawai where Kd_peg='" & Trim(DBCombo1.Text) & "'"
RS.Open p, db
If RS.BOF Or RS.EOF Then
    MsgBox "salah"
    DBCombo1.Text = ""
    Else
    Label7.Caption = Trim(RS("Nama_peg"))
    End If
End Sub

Private Sub DBCombo1_GotFocus()
tampilcombo
End Sub

Private Sub DBCombo1_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    If DBCombo1.Text = "" Then
        MsgBox "Isikan Kode Pegawai dahulu !", vbCritical + vbOKOnly, "Peringatan !"
        DBCombo1.SetFocus
        Else
        status
        Dim p
        p = "select Kd_peg from pegawai where Kd_peg='" & Trim(DBCombo1.Text) & "'"
        RS.Open p, db
        If RS.EOF Then
            MsgBox "Kode Pegawai belum terdaftar !", vbCritical + vbOKOnly, "Peringatan !"
            DBCombo1.SetFocus
            Else
            status
            Dim X
            X = "select * from pegawai where Kd_peg='" & Trim(DBCombo1.Text) & "'"
            RS.Open X, db
            If RS.BOF Or RS.EOF Then
                MsgBox "salah"
                DBCombo1.Text = ""
                Else
                Label7.Caption = Trim(RS("Nama_peg"))
                End If
            Text3.Enabled = True
            Text3.SetFocus
            End If
    End If
End If
End Sub

Private Sub Form_Load()
Set Data = OpenDatabase(App.Path & "\sepatu.mdb")
Set Data1.Recordset = Data.OpenRecordset("user")
Set Data2.Recordset = Data.OpenRecordset("pegawai")
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
menuutama.StatusBar1.Panels(1).Text = "Input Data User"
End Sub
Sub atur()
Grid.Columns(0).Caption = "Kode User"
Grid.Columns(0).Width = "1100"
Grid.Columns(1).Caption = "Kode Pegawai"
Grid.Columns(1).Alignment = dbgGeneral
Grid.Columns(1).Width = "1200"
Grid.Columns(2).Caption = "Username"
Grid.Columns(2).Width = "1200"
Grid.Columns(3).Caption = "Password"
Grid.Columns(3).Width = "1200"
Grid.Columns(4).Caption = "Status"
Grid.Columns(4).Width = "1200"
End Sub
Sub kosong()
Text1.Text = ""
DBCombo1.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
menuutama.Enabled = True
menuutama.StatusBar1.Panels(1).Text = "Halaman Utama System"
End Sub

Private Sub Grid_DblClick()
Text1.Text = Trim(RS("Kd_user"))
DBCombo1.Text = Trim(RS("Kd_peg"))
Text3.Text = Trim(RS("Username"))
Text4.Text = Trim(RS("Password"))
Text5.Text = Trim(RS("Status"))
bukakunci
btnedit.Enabled = True
btnhapus.Enabled = True
End Sub

Sub tampilcombo()
status
RS.Open "select * from pegawai", db
DBCombo1.Clear
Do While Not RS.EOF
    DBCombo1.AddItem RS!Kd_peg
    RS.MoveNext
Loop
End Sub


Private Sub Text1_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then
    If Text1.Text = "" Then
        MsgBox "Isikan Kode [user] dahulu !", vbCritical + vbOKCancel, "Peringatan"
        Else
        Dim a
    status
    a = "select * from [user] where Kd_user='" & Trim(Text1.Text) & "'"
    RS.Open a, db
    If RS.BOF Or RS.EOF Then
        DBCombo1.Enabled = True
        DBCombo1.SetFocus
        Grid.Enabled = False
        Else
        MsgBox "Kode User sudah ada !", vbCritical + vbOKOnly, "Peringatan !"
        Matikan
        Text1.Text = ""
        Text1.SetFocus
    End If
End If
End If
End Sub
Sub Matikan()
DBCombo1.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
End Sub

Private Sub Text2_Change()
If Trim(Text2) <> Empty Then
    Cari "Kd_user", Text2.Text
Else
    tampil
End If
End Sub

Private Sub Text2_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
End Sub

Private Sub Text3_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then
    If Text3.Text = "" Then
        MsgBox "Isikan Username dahulu !", vbCritical + vbOKOnly, "Peringatan !"
        Text3.SetFocus
        Else
        Text4.Enabled = True
        Text4.SetFocus
    End If
End If
End Sub

Sub bukakunci()
Text2.Enabled = True
DBCombo1.Enabled = True
Text3.Enabled = True
End Sub
Private Sub Text4_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then
    If Text4.Text = "" Then
        MsgBox "Isikan Password dahulu !", vbCritical + vbOKOnly, "Peringatan !"
        Text4.SetFocus
        Else
        Text5.Enabled = True
        Text5.SetFocus
    End If
End If
End Sub
Private Sub Text5_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then
    If Text5.Text = "" Then
        MsgBox "Isikan Status dahulu !", vbCritical + vbOKOnly, "Peringatan !"
        Text5.SetFocus
        Else
        btnsimpan.Enabled = True
    End If
End If
End Sub
Sub Cari(Kd_user, Username)
Dim p
status
p = "select * from user where " & Kd_user & " like '%" & Username & "%' order by Kd_user"
RS.Open p, db
Set Grid.DataSource = RS
Grid.Refresh
End Sub

Private Sub Text6_Change()
If Trim(Text6) <> Empty Then
    Cari "Username", Text6.Text
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
