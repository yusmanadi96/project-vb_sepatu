Attribute VB_Name = "Modbas"
Option Explicit

Public db As New ADODB.Connection
Public rsBarang As New ADODB.Recordset
Public rspenjualan As New ADODB.Recordset
Public rspeg As New ADODB.Recordset
Public RSKasir As ADODB.Recordset
Public rspembelian As ADODB.Recordset
Public RSDetailBeli As ADODB.Recordset
Public RSTransaksi As ADODB.Recordset
Public RSPemasok As ADODB.Recordset
Public rsuser As ADODB.Recordset
Public rssup As ADODB.Recordset



Sub koneksi()
     On Error GoTo CEK
    
    Set db = New ADODB.Connection
    db.CursorLocation = adUseClient
    
    Set db = New ADODB.Connection
    db.Open "provider=microsoft.jet.OLEDB.4.0;Data Source=" & App.Path & "\sepatu.mdb;Jet OLEDB:Database"
    
    
    Set rsBarang = New ADODB.Recordset
    Set RSKasir = New ADODB.Recordset
    Set rspembelian = New ADODB.Recordset
    Set RSDetailBeli = New ADODB.Recordset
    Set RSTransaksi = New ADODB.Recordset
    Set RSPemasok = New ADODB.Recordset
    Set rsuser = New ADODB.Recordset
        Set rssup = New ADODB.Recordset

    Exit Sub
CEK:
MsgBox "Koneksi Error : " & Err.Description, vbCritical, "Koneksi Error"
End
End Sub

Public Sub SetFormCenter(Frm As Form)
Frm.Move (menuutama.ScaleWidth \ 2) - (Frm.Width \ 2), (menuutama.ScaleHeight / 2) - (Frm.Height / 2)
End Sub
Public Function AwalKataKapital(strKalimat As String)
Dim i As Integer
Dim Temp As String
Dim Lokasi As Integer
Dim huruf As String * 1
  Temp$ = ""
  
  For i% = 1 To Len(strKalimat)
    huruf = Chr(Asc(Mid(strKalimat, i%, 1)))
    If Len(Trim(huruf)) < 1 Then
      Lokasi% = i% + 1
    End If
    If i% = Lokasi% Or i% = 1 Then
       Temp$ = Temp$ + UCase(Chr(Asc(Mid(strKalimat, _
               i%, 1))))
    Else
       Temp$ = Temp$ + LCase(Chr(Asc(Mid(strKalimat, _
                i%, 1))))
    End If
  Next i
  AwalKataKapital = Temp$
End Function
Public Function Cetak_Beli()
Call koneksi
'cari faktur terakhir
'status
rspembelian.Open "select * from Pengadaan Where Faktur In(Select Max(Faktur)From Pengadaan)Order By Faktur Desc", db
Dim JmlHarga, JmlBeli, JmlHasil As Double
Dim MGrs As String
Printer.Font = "Courier New"
Printer.Print
Printer.Print
'buka tabel kasir dan pemasok
rspeg.Open "select * From Pegawai where Kd_peg= '" & rspembelian!Kd_peg & "'", db
rssup.Open "select * From Supplier where Kd_sup= '" & rspembelian!Kd_sup & "'", db

'cetak data ke printer
Printer.CurrentX = 0
Printer.CurrentY = 0
Printer.Print Tab(5); "Faktur     :   "; rspembelian!Faktur
Printer.Print Tab(5); "Tanggal    :   "; Format(rspembelian!Tanggal, "DD-MMMM-YYYY")
Printer.Print Tab(5); "Jam        :   "; Format(rspembelian!Jam, "HH:MM:SS")
Printer.Print Tab(5); "Kasir      :   "; rspeg!Nama_peg
Printer.Print Tab(5); "Pemasok    :   "; rssup!Nama_sup
Printer.Print Tab(5); "Telepon    :   "; rssup!Telepon

MGrs = String$(33, "-")
Printer.Print Tab(5); MGrs

'cari data di tabel detailbeli yang fakturnya =di tbl pembelian
RSDetailBeli.Open "select * from DetailBeli Where Faktur='" & rspembelian!Faktur & "'", db
RSDetailBeli.MoveFirst
Dim no
no = 0
Do While Not RSDetailBeli.EOF
    no = no + 1
    
    Set rsBarang = New ADODB.Recordset
    'cari barang yang kodenya disimpan di tabel detailbeli
    rsBarang.Open "select * From Barang where Kodebrg= '" & RSDetailBeli!KodeBrg & "'", db
    rsBarang.Requery
    JmlHarga = rsBarang!HargaBrg
    JmlBeli = RSDetailBeli!JmlBeli
    JmlHasil = JmlHarga * JmlBeli
    'printer berulang-ulang kode,nama,harga,jumlah dan total
    Printer.Print Tab(5); no; Space(2); rsBarang!NamaBrg
    Printer.Print Tab(10); RKanan(JmlBeli, "##"); Space(1); "X";
    Printer.Print Tab(15); Format(JmlHarga, "###,###,###");
    Printer.Print Tab(25); RKanan(JmlHasil, "###,###,###")
    RSDetailBeli.MoveNext
Loop

'printer total harga
Printer.Print Tab(5); MGrs
Printer.Print Tab(5); "Total      :";
Printer.Print Tab(25); RKanan(rspembelian!JmlHarga, "###,###,###");
Printer.Print Tab(5); "Dibayar    :";
'printer dibayar
Printer.Print Tab(25); RKanan(rspembelian!Dibayar, "###,###,###");
Printer.Print Tab(5); MGrs
Printer.Print Tab(5); "Kembali    :";
'printer kembalian
If rspembelian!Dibayar = rspembelian!JmlHarga Then
    Printer.Print Tab(34); rspembelian!Dibayar - rspembelian!JmlHarga
Else
    Printer.Print Tab(25); RKanan(rspembelian!Dibayar - rspembelian!JmlHarga, "###,###,###");
End If
Printer.Print Tab(5); MGrs
Printer.Print
Printer.Print
Printer.Print
Printer.EndDoc
db.Close

End Function

Public Function RKanan(NData, CFormat) As String
    RKanan = Format(NData, CFormat)
    RKanan = Space(Len(CFormat) - Len(RKanan)) + RKanan
End Function

