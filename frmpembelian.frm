VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmpembelian 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Input Data Pembelian"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   9990
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   2280
      TabIndex        =   30
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text12 
      Height          =   375
      Left            =   4320
      TabIndex        =   28
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   4320
      TabIndex        =   27
      Top             =   910
      Width           =   2415
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   4320
      TabIndex        =   26
      Top             =   515
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   4320
      TabIndex        =   25
      Top             =   160
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   375
      Left            =   960
      TabIndex        =   11
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   375
      Left            =   960
      TabIndex        =   10
      Top             =   515
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   375
      Left            =   960
      TabIndex        =   9
      Top             =   910
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   375
      Left            =   960
      TabIndex        =   8
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Top             =   4725
      Width           =   1335
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   5130
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1800
      Top             =   3600
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   5760
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid Grid2 
      Height          =   5415
      Left            =   6840
      TabIndex        =   1
      Top             =   120
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   9551
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2880
      Top             =   3720
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
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
   Begin Project1.Button btnsimpan 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   4320
      Width           =   855
      _ExtentX        =   1508
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
      TabIndex        =   7
      Top             =   1800
      Width           =   6615
      _ExtentX        =   11668
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
   Begin Project1.Button btnbatal 
      Height          =   375
      Left            =   990
      TabIndex        =   12
      Top             =   4320
      Width           =   855
      _ExtentX        =   1508
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
   Begin Project1.Button btntutup 
      Height          =   375
      Left            =   1880
      TabIndex        =   29
      Top             =   4320
      Width           =   855
      _ExtentX        =   1508
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
      Caption         =   "Tutup"
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
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Telepon"
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
      Left            =   2880
      TabIndex        =   24
      Top             =   1320
      Width           =   765
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat"
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
      Left            =   2880
      TabIndex        =   23
      Top             =   960
      Width           =   660
   End
   Begin VB.Label Label10 
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
      Left            =   2880
      TabIndex        =   22
      Top             =   600
      Width           =   1380
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Supplier"
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
      Left            =   2880
      TabIndex        =   21
      Top             =   240
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Faktur"
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
      Left            =   120
      TabIndex        =   20
      Top             =   240
      Width           =   600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
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
      Left            =   120
      TabIndex        =   19
      Top             =   600
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jam"
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
      Left            =   120
      TabIndex        =   18
      Top             =   960
      Width           =   360
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pegawai"
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
      Left            =   120
      TabIndex        =   17
      Top             =   1320
      Width           =   780
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Item"
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
      Left            =   3000
      TabIndex        =   16
      Top             =   4320
      Width           =   465
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total "
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
      Left            =   4440
      TabIndex        =   15
      Top             =   4365
      Width           =   525
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dibayar"
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
      Left            =   4440
      TabIndex        =   14
      Top             =   4770
      Width           =   735
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kembali"
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
      Left            =   4440
      TabIndex        =   13
      Top             =   5160
      Width           =   780
   End
End
Attribute VB_Name = "frmpembelian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rssup As New ADODB.Recordset
Dim rsbar As New ADODB.Recordset
Dim rstrans As New ADODB.Recordset
Dim rspem As New ADODB.Recordset
Dim rsdet As New ADODB.Recordset
Dim rsdetbel As New ADODB.Recordset
Dim rspeg As New ADODB.Recordset

Sub status()
Set rssup = New ADODB.Recordset
rssup.CursorLocation = adUseClient
Set rsbar = New ADODB.Recordset
rsbar.CursorLocation = adUseClient
Set rstrans = New ADODB.Recordset
rstrans.CursorLocation = adUseClient
Set rspem = New ADODB.Recordset
rspem.CursorLocation = adUseClient
Set rsdet = New ADODB.Recordset
rsdet.CursorLocation = adUseClient
Set rsdetbel = New ADODB.Recordset
rsdetbel.CursorLocation = adUseClient
Set rspeg = New ADODB.Recordset
rspeg.CursorLocation = adUseClient
End Sub

Sub kosong()
Combo1.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
End Sub
Sub kunci()
Text10.Enabled = False
Text11.Enabled = False
Text12.Enabled = False
End Sub
Sub bukakunci()
Text10.Enabled = True
Text11.Enabled = True
Text12.Enabled = True
End Sub

Private Sub btnbatal_Click()
    Text5 = ""
    Text6 = ""
    Text7 = ""
    Text8 = ""
    Form_Activate
End Sub

Private Sub btntutup_Click()
Unload Me
End Sub


Private Sub Combo1_Click()
   Call koneksi
    status
    rssup.Open "Select * from Supplier where Kd_sup='" & Combo1 & "'", db
    'jika ditemukan tampilkan datanya
    If Not rssup.EOF Then
        Text10 = rssup!Nama_sup
        Text11 = rssup!Alamat
        Text12 = rssup!Telepon
    End If
    db.Close
End Sub

Private Sub Form_Activate()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\sepatu.mdb"
Adodc1.RecordSource = "Transaksi"
Set Grid.DataSource = Adodc1
Grid.Refresh

koneksi
status
rsbar.Open "select NamaBrg,KodeBrg from barang", db
Set Grid2.DataSource = rsbar
Grid2.Refresh

rssup.Open "supplier", db
Combo1.Clear
Do Until rssup.EOF
    Combo1.AddItem rssup!Kd_sup
    rssup.MoveNext
Loop
Auto
Tabel_Kosong
Adodc1.Recordset.MoveFirst
Text2.Text = Date
Text4.Text = menuutama.StatusBar1.Panels(2)
btnsimpan.Enabled = False
kosong
kunci
menuutama.StatusBar1.Panels(1).Text = "Form Pembelian Barang"

End Sub

Function Tabel_Kosong()
    Adodc1.Recordset.MoveFirst
    Do While Not Adodc1.Recordset.EOF
        Adodc1.Recordset.Delete
        Adodc1.Recordset.MoveNext
    Loop
    
    For i = 1 To 1
        Adodc1.Recordset.AddNew
        Adodc1.Recordset!Nomor = i
        Adodc1.Recordset.Update
    Next i
    Grid.Col = 1
End Function
Private Sub Auto()
Call koneksi
status
rspem.Open "select * from Pengadaan Where Faktur In(Select Max(Faktur)From Pengadaan)Order By Faktur Desc", db
rspem.Requery
    Dim Urutan As String * 10
    Dim Hitung As Long
    With rspem
        'jika tidak ditemukan maka...
        If .EOF Then
            Urutan = Right(Date, 2) + Mid(Date, 4, 2) + Left(Date, 2) + "0001"
            'no fakturnya adalah YYMMDD0001
            Text1 = Urutan
        Else
            'jika ganti hari maka... nomor fakturnya
            If Left(!Faktur, 6) <> Right(Date, 2) + Mid(Date, 4, 2) + Left(Date, 2) Then
                'YYMMDD0001
                Urutan = Right(Date, 2) + Mid(Date, 4, 2) + Left(Date, 2) + "0001"
            Else
                'jika harinya sama maka... YYMMDD0001+1
                Hitung = (!Faktur) + 1
                Urutan = (Right(Date, 2) + Mid(Date, 4, 2) + Left(Date, 2)) + Right("0000" & Hitung, 4)
            End If
        End If
        Text1 = Urutan
    End With
End Sub




Private Sub Form_Load()
Call SetFormCenter(Me)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
menuutama.Enabled = True
menuutama.StatusBar1.Panels(1).Text = "Halaman Utam System"
End Sub

Private Sub Grid_Keypress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Grid.Col = 3 Then
    'kolom 3 dan 4 hanya dapat diisi angka
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack Or Keyascii = vbKeyReturn) Then Keyascii = 0
ElseIf Grid.Col = 4 Then
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack Or Keyascii = vbKeyReturn) Then Keyascii = 0
End If
End Sub




Private Sub Text4_Change()
status
rspeg.Open "select * from [user] where Username='" & Trim(Text4.Text) & "'", db
If Not rspeg.EOF Then
    Text13 = Trim(rspeg("Kd_peg"))
    Else
    MsgBox "salah"
    End If
End Sub

Private Sub Timer1_Timer()
Text3.Text = Time$
End Sub
Function Tambah_Baris()
    For i = Adodc1.Recordset.RecordCount To Adodc1.Recordset.RecordCount
        Adodc1.Recordset.AddNew
        Adodc1.Recordset!Nomor = i + 1
        Adodc1.Recordset.Update
    Next i
End Function

Private Sub Grid_AfterColEdit(ByVal ColIndex As Integer)
    If Grid.Col = 1 Then
        'kode barang harus 6 digit
        If Len(Adodc1.Recordset!Kode) < 6 Then
            MsgBox "Kode Harus 6 digit"
            Grid.Col = 1
            Exit Sub
        End If
    
        Call koneksi
        'cari barang yg kodenya diketik di grid
        status
        rsbar.Open "Select * from Barang where Kodebrg='" & Adodc1.Recordset!Kode & "'", db
        'jika tidak ada munculkan pesan
        If rsbar.EOF Then
            MsgBox ("Ini Kode Barang baru, isi data dengan lengkap")
            Adodc1.Recordset!Kode = Adodc1.Recordset!Kode
            'isi nama barang (karena ini barang baru)
            Grid.Col = 2
            Grid.Refresh
            Exit Sub
        Else
            'jika ditemukan tampilkan nama,harga dst...
            Adodc1.Recordset!Kode = rsbar!KodeBrg
            Adodc1.Recordset!Nama = rsbar!NamaBrg
            Adodc1.Recordset!Harga = rsbar!HargaBrg
            Grid.Col = 4
            Grid.Refresh
            Exit Sub
        End If
    End If
    
    'isi nama barang jika barang baru
    If Grid.Col = 2 Then
        Adodc1.Recordset!Nama = Adodc1.Recordset!Nama
        Adodc1.Recordset.Update
        Grid.Col = 3
        Grid.Refresh
        Exit Sub
    End If
    
    'isi harga barang jika barang baru
    If Grid.Col = 3 Then
        Adodc1.Recordset!Harga = Adodc1.Recordset!Harga
        Adodc1.Recordset.Update
        Grid.Col = 4
        Grid.Refresh
        Exit Sub
    End If
    
    'isi jumlah barang jika barang baru
    If Grid.Col = 4 Then
        Adodc1.Recordset!jumlah = Adodc1.Recordset!jumlah
        'total dihasilkan dari harga x jumlah
        Adodc1.Recordset!Total = Adodc1.Recordset!Harga * Adodc1.Recordset!jumlah
        Adodc1.Recordset.Update
        Call Tambah_Baris
        Adodc1.Recordset.MoveNext
        Grid.Col = 1
        Adodc1.Recordset.MoveLast
        'tampilkan total harga dan total item
        Call TotalHarga
        Call TotalItem
    End If
End Sub
 Sub Bersihkan()
    Text5 = ""
    Text6 = ""
    Text7 = ""
    Text8 = ""
    Call kosong
End Sub
Private Sub Text7_KeyPress(Keyascii As Integer)
    If Keyascii = 13 Then
        'pembayaran tidak boleh kosong atau lebih kecil
        If Text7 = "" Or val(Text7) < (Text6) Then
            MsgBox "Jumlah Pembayaran Kurang"
            Text7.SetFocus
        Else
            Text7 = Format(Text7, "###,###,###")
            If Text7 = Text6 Then
                Text8 = Text7 - Text6
            Else
                Text8 = Format(Text7 - Text6, "###,###,###")
            End If
        btnsimpan.Enabled = True
        End If
    End If
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

Private Sub btnSimpan_Click()
If Combo1 = "" Or Text10 = "" Or Text11 = "" Or Text12 = "" Or Text7 = "" Then
    MsgBox "Data belum lengkap"
    Exit Sub
Else
    If Text5 = "" Then
        MsgBox "Tidak ada transaksi pembelian"
        Exit Sub
    End If
End If

    Call koneksi
    status
    rssup.Open "select * from Supplier where Kd_sup='" & Combo1 & "'", db
    If rssup.EOF Then
        Dim TambahPemasok As String
        TambahPemasok = "Insert Into Supplier(Kd_sup,Nama_sup,Alamat,Telepon)" & _
        "values('" & Combo1 & "','" & Text10 & "','" & Text11 & "','" & Text12 & "')"
        db.Execute (TambahPemasok)
    End If
    
    'simpan transaksi ke tbl pembelian
    Dim SQLTambahJual As String
    SQLTambahJual = "Insert Into Pengadaan(Faktur,Tanggal,Jam,JmlItem,JmlHarga,Dibayar,JmlKembali,Kd_peg,Kd_sup)" & _
    "values('" & Text1 & "','" & Text2 & "','" & Text3 & "','" & Text5 & "','" & Text6 & "','" & Text7 & "','" & Text8 & "','" & Text13 & "','" & Combo1 & "')"
    db.Execute (SQLTambahJual)
    
    'simpan data transaksi ke tabel detailbeli
    'jika ada kode yang sama maka jumlahnya akan disatukan
    status
    rstrans.Open "select kode as KodeBrg,sum(Jumlah) as JumlahBrg from Transaksi group by kode", db
    rstrans.MoveFirst
    Do While Not rstrans.EOF
        If rstrans!KodeBrg <> vbNullString Then
            Dim SQLTambahDetail As String
            SQLTambahDetail = "Insert Into DetailBeli(Faktur,Kodebrg,JmlBeli) " & _
            "values ('" & Text1 & "','" & rstrans!KodeBrg & "','" & rstrans!JumlahBrg & "')"
            db.Execute (SQLTambahDetail)
        End If
    rstrans.MoveNext
    Loop
        
    Adodc1.Recordset.MoveFirst
    Do While Not Adodc1.Recordset.EOF
        If Adodc1.Recordset!Kode <> vbNullString Then
            Call koneksi
            status
            rsbar.Open "Select * from Barang where Kodebrg='" & Adodc1.Recordset!Kode & "'", db
            If Not rsbar.EOF Then
                'tambah barang jika kodenya ditemukan
                Dim TambahBarang1 As String
                TambahBarang1 = "update barang set jumlahbrg='" & rsbar!JumlahBrg + Adodc1.Recordset!jumlah & "' where kodebrg='" & Adodc1.Recordset!Kode & "'"
                db.Execute (TambahBarang1)
            Else
                'input data barang jika kodenya baru
                Dim TambahBarang2 As String
                TambahBarang2 = "Insert Into Barang(Kodebrg,NamaBrg,HargaBrg,JumlahBrg,Kd_sup)" & _
                "values('" & Adodc1.Recordset!Kode & "','" & Adodc1.Recordset!Nama & "','" & Adodc1.Recordset!Harga & "','" & Adodc1.Recordset!Harga & "','" & Adodc1.Recordset!jumlah & "','" & Trim(Combo1) & "')"
                db.Execute (TambahBarang2)
            End If
        End If
    Adodc1.Recordset.MoveNext
    Loop
    
    Bersihkan
    Form_Activate
    Combo1.SetFocus
    'Call Cetak_Beli
End Sub
Function TotalItem()
On Error Resume Next
Adodc1.Recordset.MoveFirst
Text5 = 0
Do While Not Adodc1.Recordset.EOF And Adodc1.Recordset!jumlah <> 0
    Text5 = Text5 + Adodc1.Recordset!jumlah
    Adodc1.Recordset.MoveNext
    Text5 = Text5
Loop
End Function

Function TotalHarga()
On Error Resume Next
Adodc1.Recordset.MoveFirst
Text6 = 0
Do While Not Adodc1.Recordset.EOF And Adodc1.Recordset!Total <> 0
    Text6 = Text6 + Adodc1.Recordset!Total
    Adodc1.Recordset.MoveNext
    Text6 = Format(Text6, "#,###,###")
Loop
End Function
Function Cetak()
Call koneksi
'cari faktur terakhir
status
rspem.Open "select * from Pengadaan Where Faktur In(Select Max(Faktur)From Pengadaan)Order By Faktur Desc", db
Tampilkan.Show

Dim JmlHarga, JmlBeli, JmlHasil As Double
Dim MGrs As String
Tampilkan.Font = "Courier New"
Tampilkan.Print
Tampilkan.Print
'buka tabel kasir dan pemasok
rspeg.Open "select * From Pegawai where Kd_peg= '" & rspem!Kd_peg & "'", db
rssup.Open "select * From Supplier where Kd_sup= '" & rspem!Kd_sup & "'", db

'cetak data ke layar
Tampilkan.Print Tab(5); "Faktur     :   "; rspem!Faktur
Tampilkan.Print Tab(5); "Tanggal    :   "; Format(rspem!Tanggal, "DD-MMMM-YYYY")
Tampilkan.Print Tab(5); "Jam        :   "; Format(rspem!Jam, "HH:MM:SS")
Tampilkan.Print Tab(5); "Kasir      :   "; rspeg!Nama_peg
Tampilkan.Print Tab(5); "Pemasok    :   "; rssup!Nama_sup
Tampilkan.Print Tab(5); "Telepon    :   "; rssup!Telepon

MGrs = String$(33, "-")
Tampilkan.Print Tab(5); MGrs

'cari data di tabel detailbeli yang fakturnya =di tbl pembelian
rsdet.Open "select * from DetailBeli Where Faktur='" & rspem!Faktur & "'", db
rsdet.MoveFirst

no = 0
Do While Not rsdet.EOF
    no = no + 1
    
    Set rsbar = New ADODB.Recordset
    'cari barang yang kodenya disimpan di tabel detailbeli
    rsbar.Open "select * From Barang where Kodebrg= '" & rsdet!KodeBrg & "'", db
    rsbar.Requery
    JmlHarga = rsbar!HargaBrg
    JmlBeli = rsdet!JmlBeli
    JmlHasil = JmlHarga * JmlBeli
    'tampilkan berulang-ulang kode,nama,harga,jumlah dan total
    Tampilkan.Print Tab(5); no; Space(2); rsbar!NamaBrg
    Tampilkan.Print Tab(10); RKanan(JmlBeli, "##"); Space(1); "X";
    Tampilkan.Print Tab(15); Format(JmlHarga, "###,###,###");
    Tampilkan.Print Tab(25); RKanan(JmlHasil, "###,###,###")
    rsdet.MoveNext
Loop

'tampilkan total harga
Tampilkan.Print Tab(5); MGrs
Tampilkan.Print Tab(5); "Total      :";
Tampilkan.Print Tab(25); RKanan(rspem!JmlHarga, "###,###,###");
Tampilkan.Print Tab(5); "Dibayar    :";
'tampilkan dibayar
Tampilkan.Print Tab(25); RKanan(rspem!Dibayar, "###,###,###");
Tampilkan.Print Tab(5); MGrs
Tampilkan.Print Tab(5); "Kembali    :";
'tampilkan kembalian
If rspem!Dibayar = rspem!JmlHarga Then
    Tampilkan.Print Tab(34); rspem!Dibayar - rspem!JmlHarga
Else
    Tampilkan.Print Tab(25); RKanan(rspem!Dibayar - rspem!JmlHarga, "###,###,###");
End If
Tampilkan.Print Tab(5); MGrs
Tampilkan.Print
Tampilkan.Print
Tampilkan.Print
db.Close
End Function


Private Function RKanan(NData, CFormat) As String
    RKanan = Format(NData, CFormat)
    RKanan = Space(Len(CFormat) - Len(RKanan)) + RKanan
End Function

