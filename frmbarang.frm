VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmbarang 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Input Data Barang"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   150
      Left            =   4560
      Top             =   120
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   5040
      TabIndex        =   18
      Top             =   2760
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1800
      TabIndex        =   17
      Top             =   2760
      Width           =   1215
   End
   Begin VB.ComboBox DBCombo1 
      Height          =   315
      Left            =   2040
      TabIndex        =   12
      Top             =   960
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid Grid 
      Height          =   2295
      Left            =   120
      TabIndex        =   11
      Top             =   3360
      Width           =   7215
      _ExtentX        =   12726
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
   Begin Project1.Button btnhapus 
      Height          =   345
      Left            =   4575
      TabIndex        =   10
      Top             =   1995
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   609
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
   Begin Project1.Button btnedit 
      Height          =   345
      Left            =   3690
      TabIndex        =   9
      Top             =   1995
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   609
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
   Begin Project1.Button btnbatal 
      Height          =   345
      Left            =   2805
      TabIndex        =   8
      Top             =   1995
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   609
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
   Begin Project1.Button btnsimpan 
      Height          =   345
      Left            =   1920
      TabIndex        =   7
      Top             =   1995
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   609
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
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   2040
      TabIndex        =   6
      Top             =   1400
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   2040
      TabIndex        =   5
      Top             =   600
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   2040
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Barang"
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
      Left            =   3360
      TabIndex        =   19
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Barang"
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
      TabIndex        =   16
      Top             =   2760
      Width           =   1365
   End
   Begin VB.Label Label7 
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
      Left            =   3120
      TabIndex        =   15
      Top             =   2520
      Width           =   1200
   End
   Begin VB.Shape Shape2 
      Height          =   735
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   7215
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Supplier"
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
      Left            =   3840
      TabIndex        =   14
      Top             =   1000
      Width           =   1380
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   5400
      TabIndex        =   13
      Top             =   960
      Width           =   1785
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   495
      Left            =   1800
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   3735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Supplier"
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
      TabIndex        =   3
      Top             =   960
      Width           =   1620
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Barang"
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
      TabIndex        =   2
      Top             =   600
      Width           =   1515
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Harga Satuan"
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
      TabIndex        =   1
      Top             =   1320
      Width           =   1515
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Barang"
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
      TabIndex        =   0
      Top             =   240
      Width           =   1440
   End
End
Attribute VB_Name = "frmbarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim strTemp, LenTemp, n, enTemp

Sub tampil()
Dim p
status
p = "select * from barang"
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
Label5.Caption = ""
End Sub

Private Sub btnedit_Click()
If Text1.Text = "" Or Text2.Text = "" Or DBCombo1.Text = "" Or Text3.Text = "" Then
    MsgBox "Isikan data secara lengkap !", vbCritical + vbOKOnly, "Peringatan !"
    Text1.SetFocus
    Else
Dim p
Select Case MsgBox("Apakah data sudah benar ?", vbYesNo, "Perhatian")
Case vbYes
    status
    p = "update  barang set KodeBrg='" & Trim(Text1.Text) & "',NamaBrg='" & Trim(Text2.Text) & "',Kode_sup='" & Trim(DBCombo1.Text) & "',HargaBrg='" & Trim(Text3.Text) & "',JumlahBrg=(0) where KodeBrg='" & Trim(Grid.Columns(0)) & "'"
        
    RS.Open p, db
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
If Text1.Text = "" Or Text2.Text = "" Or DBCombo1.Text = "" Or Text3.Text = "" Then
    MsgBox "Pilih Data dahulu sebelum menghapus !", vbCritical + vbOKOnly, "Peringatan !"
    Else
Dim p
Select Case MsgBox("Apakah data akan dihapus ?", vbYesNo, "Perhatian")
Case vbYes
    status
    p = "delete * from barang where KodeBrg='" & Trim(Text1.Text) & "'"
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
If Text1.Text = "" Or Text2.Text = "" Or DBCombo1.Text = "" Or Text3.Text = "" Then
    MsgBox "Isikan data secara lengkap !", vbCritical + vbOKOnly, "Peringatan !"
    Text1.SetFocus
    Else
    Dim X
    status
    X = "insert into barang(KodeBrg,NamaBrg,Kode_sup,HargaBrg,JumlahBrg) values ('" & Trim(Text1.Text) & "','" & Trim(Text2.Text) & "','" & Trim(DBCombo1.Text) & "','" & Trim(Text3.Text) & "',0)"
    RS.Open X, db
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
p = "select * from Supplier where Kd_sup='" & Trim(DBCombo1.Text) & "'"
RS.Open p, db
If RS.BOF Or RS.EOF Then
    MsgBox "salah"
    DBCombo1.Text = ""
    Else
    Label5.Caption = Trim(RS("Nama_sup"))
    End If
End Sub

Private Sub DBCombo1_GotFocus()
tampilcombo
End Sub

Private Sub DBCombo1_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    If DBCombo1.Text = "" Then
        MsgBox "Isikan Kode Supplier dahulu !", vbCritical + vbOKOnly, "Peringatan !"
        DBCombo1.SetFocus
        Else
        status
        Dim p
        p = "select kd_sup from Supplier where kd_sup='" & Trim(DBCombo1.Text) & "'"
        RS.Open p, db
        If RS.EOF Then
            MsgBox "Kode Supplier belum terdaftar !", vbCritical + vbOKOnly, "Peringatan !"
            DBCombo1.SetFocus
            Else
             status
                Dim z
                z = "select * from Supplier where Kd_sup='" & Trim(DBCombo1.Text) & "'"
                RS.Open z, db
                If RS.BOF Or RS.EOF Then
                MsgBox "salah"
                DBCombo1.Text = ""
                Else
                Label5.Caption = Trim(RS("Nama_sup"))
                End If

            Text3.Enabled = True
            Text3.SetFocus
            End If
    End If
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
menuutama.StatusBar1.Panels(1).Text = "Input Data Barang"
End Sub
Sub atur()
Grid.Columns(0).Caption = "Kode Barang"
Grid.Columns(0).Width = "1100"
Grid.Columns(1).Caption = "Nama Barang"
Grid.Columns(1).Alignment = dbgGeneral
Grid.Columns(1).Width = "3000"
Grid.Columns(2).Caption = "Harga Barang"
Grid.Columns(2).Width = "1200"
Grid.Columns(3).Caption = "Kode Sup"
Grid.Columns(3).Width = "700"
Grid.Columns(4).Caption = "Stock"
Grid.Columns(4).Width = "1000"
End Sub
Sub kosong()
Text1.Text = ""
Text2.Text = ""
DBCombo1.Text = ""
Text3.Text = ""
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
menuutama.Enabled = True
menuutama.StatusBar1.Panels(1).Text = "Halaman Utama System"
End Sub

Private Sub Grid_DblClick()
Text1.Text = Trim(RS("KodeBrg"))
Text2.Text = Trim(RS("NamaBrg"))
DBCombo1.Text = Trim(RS("Kode_sup"))
Text3.Text = Trim(RS("HargaBrg"))
bukakunci
btnedit.Enabled = True
btnhapus.Enabled = True
End Sub

Sub tampilcombo()
status
RS.Open "select * from Supplier", db
DBCombo1.Clear
Do While Not RS.EOF
    DBCombo1.AddItem RS!Kd_sup
    RS.MoveNext
Loop
End Sub
Private Sub Text1_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then
    If Text1.Text = "" Then
        MsgBox "Isikan Kode Barang dahulu !", vbCritical + vbOKCancel, "Peringatan"
        Else
    Dim p
    status
    p = "select * from barang where KodeBrg='" & Trim(Text1.Text) & "'"
    RS.Open p, db
    If RS.BOF Or RS.EOF Then
        Text2.Enabled = True
        Text2.SetFocus
        Grid.Enabled = False
        Else
        MsgBox "Kode Barang sudah ada !", vbCritical + vbOKOnly, "Peringatan !"
        Matikan
        Text1.Text = ""
        Text1.SetFocus
    End If
End If
End If
End Sub
Sub Matikan()
Text2.Enabled = False
DBCombo1.Enabled = False
Text3.Enabled = False
End Sub
Private Sub Text2_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then
    If Text2.Text = "" Then
        MsgBox "Isikan Nama Barang dahulu !", vbCritical + vbOKOnly, "Peringatan !"
        Text2.SetFocus
        Else
        DBCombo1.Enabled = True
        DBCombo1.SetFocus
    End If
End If
End Sub

Sub bukakunci()
Text2.Enabled = True
DBCombo1.Enabled = True
Text3.Enabled = True
End Sub
Private Sub Text3_KeyPress(Keyascii As Integer)
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
Sub Cari(KodeBrg, NamaBrg)
Dim p
status
p = "select * from barang where " & KodeBrg & " like '%" & NamaBrg & "%' order by KodeBrg"
RS.Open p, db
Set Grid.DataSource = RS
Grid.Refresh
End Sub

Private Sub Text4_Change()
If Trim(Text4) <> Empty Then
    Cari "KodeBrg", Text4.Text
Else
    tampil
End If
End Sub

Private Sub Text4_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
End Sub

Private Sub Text5_Change()
If Trim(Text5) <> Empty Then
    Cari "NamaBrg", Text5.Text
Else
    tampil
End If
End Sub

Private Sub Text5_KeyPress(Keyascii As Integer)
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
