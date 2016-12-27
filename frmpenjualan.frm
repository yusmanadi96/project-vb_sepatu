VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmpenjualan 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Input Data Penjualan"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   2880
      TabIndex        =   20
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid Grid2 
      Height          =   4455
      Left            =   6240
      TabIndex        =   19
      Top             =   120
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   7858
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
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1800
      Top             =   4080
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2280
      Top             =   120
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
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4800
      TabIndex        =   18
      Top             =   4290
      Width           =   1335
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   4800
      TabIndex        =   17
      Top             =   3890
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4800
      TabIndex        =   16
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   375
      Left            =   3000
      TabIndex        =   12
      Top             =   3480
      Width           =   615
   End
   Begin Project1.Button Button1 
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3480
      Width           =   855
      _extentx        =   1508
      _extenty        =   661
      font            =   "frmpenjualan.frx":0000
      caption         =   "Simpan"
      captionhighlitecolor=   0
      iconhighlitecolor=   0
      checked         =   0   'False
      colorbuttonhover=   16760976
      colorbuttonup   =   15309136
      colorbuttondown =   15309136
      colorbright     =   16772528
      borderbrightness=   0
      displayhand     =   0   'False
      colorscheme     =   0
   End
   Begin MSDataGridLib.DataGrid Grid 
      Height          =   2415
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   6015
      _ExtentX        =   10610
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
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4920
      TabIndex        =   7
      Top             =   530
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4920
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   530
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin Project1.Button Button2 
      Height          =   375
      Left            =   1080
      TabIndex        =   10
      Top             =   3480
      Width           =   855
      _extentx        =   1508
      _extenty        =   661
      font            =   "frmpenjualan.frx":0028
      caption         =   "Batal"
      captionhighlitecolor=   0
      iconhighlitecolor=   0
      checked         =   0   'False
      colorbuttonhover=   16760976
      colorbuttonup   =   15309136
      colorbuttondown =   15309136
      colorbright     =   16772528
      borderbrightness=   0
      displayhand     =   0   'False
      colorscheme     =   0
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
      Left            =   3840
      TabIndex        =   15
      Top             =   4320
      Width           =   780
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
      Left            =   3840
      TabIndex        =   14
      Top             =   3930
      Width           =   735
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
      Left            =   3840
      TabIndex        =   13
      Top             =   3530
      Width           =   525
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Item"
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
      Left            =   2040
      TabIndex        =   11
      Top             =   3480
      Width           =   945
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
      Left            =   4080
      TabIndex        =   3
      Top             =   600
      Width           =   780
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
      Left            =   4080
      TabIndex        =   2
      Top             =   240
      Width           =   360
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
      TabIndex        =   1
      Top             =   600
      Width           =   765
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
      TabIndex        =   0
      Top             =   240
      Width           =   600
   End
End
Attribute VB_Name = "frmpenjualan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub status()
Set rsBarang = New ADODB.Recordset
rsBarang.CursorLocation = adUseClient
Set rspenjualan = New ADODB.Recordset
rspenjualan.CursorLocation = adUseClient
Set rspeg = New ADODB.Recordset
rspeg.CursorLocation = adUseClient
End Sub





Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
menuutama.Enabled = True
menuutama.StatusBar1.Panels(1).Text = "Halaman Utama System"
End Sub

Private Sub Grid2_DblClick()
status
Dim p
p = "select KodeBrg,NamaBrg,HargaBrg from barang where KodeBrg='" & Trim(Grid2.Columns(0)) & "'"
rsBarang.Open p, db
If Not rsBarang.EOF Then
    Adodc1.Recordset!Kode = rsBarang!KodeBrg
        Adodc1.Recordset!Nama = rsBarang!NamaBrg
        Adodc1.Recordset!Harga = rsBarang!HargaBrg * 1.1
        Grid.Col = 4
        Grid.Refresh
        End If
End Sub


Private Sub Text4_Change()
Dim p
status
p = "select * from [user] where Username='" & Trim(Text4.Text) & "'"
rspeg.Open p, db
Text9.Text = rspeg!Kd_peg
End Sub

Private Sub Timer1_Timer()
    Text3.Text = Time$
End Sub

Private Sub Form_Activate()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\sepatu.mdb"
Adodc1.RecordSource = "Transaksi"
Set Grid.DataSource = Adodc1
Grid.Refresh

Call koneksi
status
rsBarang.Open "select KodeBrg,NamaBrg from barang", db
Set Grid2.DataSource = rsBarang
Grid.Refresh
aturgrid2

Text4.Text = menuutama.StatusBar1.Panels(2)

Call Auto
Call Tabel_Kosong
Adodc1.Recordset.MoveFirst

Text2.Text = Date
Button1.Enabled = False
menuutama.StatusBar1.Panels(1).Text = "Form Penjualan Barang"
End Sub

Private Sub Form_Load()
    Grid.Col = 1
    Button1.Enabled = False
    Call SetFormCenter(Me)
End Sub

Private Sub Auto()
Call koneksi
status
rspenjualan.Open "select * from Penjualan Where Faktur In(Select Max(Faktur)From Penjualan)Order By Faktur Desc", db
rspenjualan.Requery
    Dim Urutan As String * 10
    Dim Hitung As Long
    With rspenjualan
        If .EOF Then
            Urutan = Right(Date, 2) + Mid(Date, 4, 2) + Left(Date, 2) + "0001"
            Text1 = Urutan
        Else
            If Left(!Faktur, 6) <> Right(Date, 2) + Mid(Date, 4, 2) + Left(Date, 2) Then
                Urutan = Right(Date, 2) + Mid(Date, 4, 2) + Left(Date, 2) + "0001"
            Else
                Hitung = (!Faktur) + 1
                Urutan = (Right(Date, 2) + Mid(Date, 4, 2) + Left(Date, 2)) + Right("0000" & Hitung, 4)
            End If
        End If
        Text1 = Urutan
    End With
    
End Sub

Function Tabel_Kosong()
If Adodc1.Recordset.RecordCount = 0 Then
    Adodc1.Recordset.AddNew
    Adodc1.Recordset!Nomor = 1
    Adodc1.Recordset.Update
    Adodc1.Recordset.MoveFirst
    Else
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
End If
End Function

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyEscape
        Adodc1.Recordset!Kode = Null
        Adodc1.Recordset!Nama = Null
        Adodc1.Recordset!Harga = Null
        Adodc1.Recordset!jumlah = Null
        Adodc1.Recordset!Total = Null
        Adodc1.Recordset.Update
        Call TotalItem
        Call TotalHarga
        Grid.Refresh
End Select
End Sub

Private Sub Grid_AfterColEdit(ByVal ColIndex As Integer)
    If Grid.Col = 1 Then
        Call koneksi
        status
        rsBarang.Open "Select * from Barang where Kodebrg='" & Adodc1.Recordset!Kode & "'", db
        If rsBarang.EOF Then
            Dim pesan
            pesan = MsgBox("Kode Barang Tidak Terdaftar")
            Grid.Col = 1
            Exit Sub
        End If
        Adodc1.Recordset!Kode = rsBarang!KodeBrg
        Adodc1.Recordset!Nama = rsBarang!NamaBrg
        Adodc1.Recordset!Harga = rsBarang!HargaBrg * 1.1
        'Adodc1.Recordset!Jumlah = rsBarang!JumlahBrg
        Grid.Col = 4
        Grid.Refresh
        Exit Sub
    End If
    
    If Grid.Col = 4 Then
        Call koneksi
        status
        rsBarang.Open "Select JumlahBrg from Barang where Kodebrg='" & Trim(Grid.Columns(1)) & "'", db
        If rsBarang!JumlahBrg < val(Grid.Columns(4)) Then
        MsgBox "Stok barang tidak mencukupi, Periksa inputan mungkin ada kesalahan", vbCritical + vbOKOnly, "Peringatan"
        Else
        Adodc1.Recordset!jumlah = Adodc1.Recordset!jumlah
        Adodc1.Recordset!Total = Adodc1.Recordset!Harga * Adodc1.Recordset!jumlah
        Adodc1.Recordset.Update
         Call Tambah_Baris
        Adodc1.Recordset.MoveNext
        Grid.Col = 1
        Adodc1.Recordset.MoveLast
        Call TotalHarga
        Call TotalItem
        Grid.Refresh
        End If
    End If
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

Private Sub Bersihkan()
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
End Sub

Private Sub Text7_KeyPress(Keyascii As Integer)
    If Keyascii = 13 Then
        If Text7.Text = "" Or val(Text7.Text) < (Text6.Text) Then
            MsgBox "Jumlah Pembayaran Kurang"
            Text7.SetFocus
        Else
            Text7.Text = Format(Text7, "###,###,###")
            If Text7.Text = Text6.Text Then
                Text8.Text = Text7.Text - Text6.Text
            Else
                Text8.Text = Format(Text7 - Text6, "###,###,###")
            End If
        Button1.Enabled = True
        End If
    End If
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

Private Sub Button1_Keypress(Keyascii As Integer)
    If Keyascii = 27 Then
        Button1.Enabled = False
        Text7.Text = ""
        Text7.SetFocus
    End If
End Sub

Private Sub Button1_Click()
   
    Dim SQLTambahJual As String
    SQLTambahJual = "Insert Into Penjualan(Faktur,Tanggal,Jam,JmlHarga,JmlItem,Dibayar,JmlKembali,Kd_peg)" & _
    "values('" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & Text6.Text & "','" & Text5.Text & "','" & Text7.Text & "','" & Text8.Text & "','" & Text9.Text & "')"
    db.Execute (SQLTambahJual)
         
    Adodc1.Recordset.MoveFirst
    Do While Not Adodc1.Recordset.EOF
        If Adodc1.Recordset!Kode <> vbNullString Then
            Dim SQLTambahDetail As String
            SQLTambahDetail = "Insert Into Detailjual(DetId,Faktur,Kodebrg,JmlJual) " & _
            "values ('" & Text1 + Adodc1.Recordset!Nomor & "','" & Text1 & "','" & Adodc1.Recordset!Kode & "','" & Adodc1.Recordset!jumlah & "')"
            db.Execute (SQLTambahDetail)
        End If
    Adodc1.Recordset.MoveNext
    Loop
        
    Adodc1.Recordset.MoveFirst
    Do While Not Adodc1.Recordset.EOF
        If Adodc1.Recordset!Kode <> vbNullString Then
            Call koneksi
            status
            rsBarang.Open "Select * from Barang where Kodebrg='" & Adodc1.Recordset!Kode & "'", db
            If Not rsBarang.EOF Then
                Dim Kurangi As String
                Kurangi = "update barang set jumlahbrg='" & rsBarang!JumlahBrg - Adodc1.Recordset!jumlah & "' where kodebrg='" & Adodc1.Recordset!Kode & "'"
                db.Execute (Kurangi)
                End If
            'End If
        End If
    Adodc1.Recordset.MoveNext
    Loop
    Bersihkan
    Form_Activate
    'Call Cetak
End Sub

Private Sub Button2_Click()
    Text7.Text = ""
    Text6.Text = ""
    Text5.Text = ""
    Form_Activate
End Sub
'Function Cetak()
'Call BukaDB
'RSPenjualan.Open "select * from penjualan Where Faktur In(Select Max(Faktur)From penjualan)Order By Faktur Desc", Conn
'Layar.Show
'Dim Total, JmlJual, JmlHasil As Double
'Dim MGrs As String
'Layar.Font = "Courier New"
'Layar.Print
'Layar.Print
'RSkasir.Open "select * From Kasir where KodeKsr= '" & RSPenjualan!KodeKsr & "'", Conn
'Layar.Print Tab(5); "Faktur     :   "; RSPenjualan!Faktur
'Layar.Print Tab(5); "Tanggal    :   "; Format(RSPenjualan!Tanggal, "DD-MMMM-YYYY")
'Layar.Print Tab(5); "Jam        :   "; Format(RSPenjualan!Jam, "HH:MM:SS")
'Layar.Print Tab(5); "Kasir      :   "; RSkasir!NamaKsr
'MGrs = String$(33, "-")
'Layar.Print Tab(5); MGrs
'RSDetailJual.Open "select * from detailjual Where left(Faktur,10)='" & RSPenjualan!Faktur & "'", Conn
'RSDetailJual.MoveFirst
'No = 0
'Do While Not RSDetailJual.EOF
'    No = No + 1
'    Set RSBarang = New ADODB.Recordset
'    RSBarang.Open "select * From Barang where Kodebrg= '" & RSDetailJual!KodeBrg & "'", Conn
'    RSBarang.Requery
'    Harga = RSBarang!HargaJual
'    Jumlah = RSDetailJual!JmlJual
'    Hasil = Harga * Jumlah
'    Layar.Print Tab(5); No; Space(2); RSBarang!NamaBrg
'    Layar.Print Tab(10); RKanan(Jumlah, "##"); Space(1); "X";
'    Layar.Print Tab(15); Format(Harga, "###,###,###");
'    Layar.Print Tab(25); RKanan(Hasil, "###,###,###")
'    RSDetailJual.MoveNext
'Loop
'Layar.Print Tab(5); MGrs
'Layar.Print Tab(5); "Total      :";
'Layar.Print Tab(25); RKanan(RSPenjualan!Total, "###,###,###");
'Layar.Print Tab(5); "Dibayar    :";
'Layar.Print Tab(25); RKanan(RSPenjualan!Dibayar, "###,###,###");
'Layar.Print Tab(5); MGrs
'Layar.Print Tab(5); "Kembali    :";
'If RSPenjualan!Dibayar = RSPenjualan!Total Then
'    Layar.Print Tab(34); RSPenjualan!Dibayar - RSPenjualan!Total
'Else
'    Layar.Print Tab(25); RKanan(RSPenjualan!Dibayar - RSPenjualan!Total, "###,###,###");
'End If
'Layar.Print Tab(5); MGrs
'Layar.Print Tab(5); "Terima Kasih atas kunjungan Anda"
'Layar.Print
'Layar.Print
'Layar.Print
'Conn.Close
'End Function

Private Function RKanan(NData, CFormat) As String
    RKanan = Format(NData, CFormat)
    RKanan = Space(Len(CFormat) - Len(RKanan)) + RKanan
End Function

Function Tambah_Baris()

   For i = Adodc1.Recordset.RecordCount To Adodc1.Recordset.RecordCount
        Adodc1.Recordset.AddNew
        Adodc1.Recordset!Nomor = i + 1
        Adodc1.Recordset.Update
    Next i
End Function
Sub aturgrid2()
Grid2.Columns(0).Caption = "Kode Barang"
Grid2.Columns(0).Width = "1100"
Grid2.Columns(1).Caption = "Nama Barang"
Grid2.Columns(1).Width = "3000"
End Sub

