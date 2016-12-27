VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm menuutama 
   BackColor       =   &H8000000C&
   Caption         =   "System Persediaan Barang Toko Sepatu"
   ClientHeight    =   8160
   ClientLeft      =   225
   ClientTop       =   525
   ClientWidth     =   15120
   LinkTopic       =   "MDIForm1"
   Picture         =   "menuutama.frx":0000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   150
      Left            =   120
      Top             =   120
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7785
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8819
            MinWidth        =   8819
            Text            =   "Halaman Utama System"
            TextSave        =   "Halaman Utama System"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnilog 
         Caption         =   "Log Out"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnikel 
         Caption         =   "Keluar"
      End
   End
   Begin VB.Menu mnuentry 
      Caption         =   "Entry Data"
      Begin VB.Menu mnidatbar 
         Caption         =   "Data Barang"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnidatsup 
         Caption         =   "Data Supplier"
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnidatpeg 
         Caption         =   "Data Pegawai"
      End
   End
   Begin VB.Menu mnutrans 
      Caption         =   "Transaksi"
      Begin VB.Menu mnitranspemb 
         Caption         =   "Transaksi Pembelian"
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnitranspenj 
         Caption         =   "Transaksi Penjualan"
      End
   End
   Begin VB.Menu mnulap 
      Caption         =   "Laporan"
      Begin VB.Menu mnilapbar 
         Caption         =   "Laporan Data Barang"
      End
      Begin VB.Menu sep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnilapsup 
         Caption         =   "Laporan Data Supplier"
      End
      Begin VB.Menu sep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnilappeg 
         Caption         =   "Laporan Data Pegawai"
      End
      Begin VB.Menu sep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnilaptrans 
         Caption         =   "Laporan Transaksi"
         Begin VB.Menu lappemb 
            Caption         =   "Pembelian"
         End
         Begin VB.Menu sep8 
            Caption         =   "-"
         End
         Begin VB.Menu lappenj 
            Caption         =   "Penjualan"
         End
      End
   End
   Begin VB.Menu mnuopt 
      Caption         =   "Option"
      Begin VB.Menu mniuser 
         Caption         =   "Input Data User"
      End
   End
End
Attribute VB_Name = "menuutama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strTemp, LenTemp, n, enTemp


Private Sub lappemb_Click()
rptpembelian.Show
End Sub

Private Sub lappenj_Click()
rptpenjualan.Show
End Sub

Private Sub MDIForm_Load()
strTemp = Me.Caption
    n = 1

End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Jawab As Integer
   Jawab = MsgBox("Anda yakin akan keluar dari program?", vbQuestion + vbYesNo, "Konfirmasi Keluar")
   Unload frmlogin
    If Jawab = vbNo Then Cancel = -1
End Sub

Private Sub mnidatbar_Click()
frmbarang.Show
Me.Enabled = False
End Sub

Private Sub mnidatpeg_Click()
frmpegawai.Show
Me.Enabled = False
End Sub

Private Sub mnidatsup_Click()
frmsupplier.Show
Me.Enabled = False
End Sub

Private Sub mnikel_Click()
Dim Form As Form               '(Sebenarnya, hal ini 'sama dengan 'End')
   For Each Form In Forms
       Unload Form
       Set Form = Nothing      'Bersihkan memori yang digunakan sebelumnya
   Next Form
Unload Me
End Sub

Private Sub mnilapbar_Click()
rptbarang.Show
End Sub

Private Sub mnilappeg_Click()
rptpegawai.Show
End Sub

Private Sub mnilapsup_Click()
rptsupplier.Show
End Sub

Private Sub mnilog_Click()
Me.Hide
frmlogin.Show
End Sub

Private Sub mnitranspemb_Click()
frmpembelian.Show
Me.Enabled = False
End Sub

Private Sub mnitranspenj_Click()
frmpenjualan.Show
Me.Enabled = False
End Sub

Private Sub mniuser_Click()
frmuser.Show
Me.Enabled = False
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
StatusBar1.Panels(4).Text = Format(Time, "hh :mm :ss") + " WIB"
StatusBar1.Panels(5).Text = Format(Date, "dd:mmmm:yyyy")
StatusBar1.Panels(6).Text = Format(Now, "dddd")
End Sub
