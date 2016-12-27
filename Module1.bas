Attribute VB_Name = "Module1"
Option Explicit

Public Conn As New ADODB.Connection
Public rsRS As New ADODB.Recordset

 

Public AddFlag As Boolean
Public EditFlag As Boolean
Public Isitext As String
Public List As ListItem
Public i As Integer
Public CariItem
Public txt As Control
Public reply As String
Public StrSql As String
Public SQlSimpan As String
Public SQLHapus As String
Public SQLPerbaiki As String

Public Sub Connect()
 
 
    Set Conn = New ADODB.Connection
    Conn.ConnectionString = "Provider =Microsoft.Jet.OLEDB.4.0;" & _
                          "Data Source=" & App.Path & "\sepatu.mdb"
    Conn.Open

 
End Sub
 
 
Public Sub LoadDataToListView(StrSql As String, rs As ADODB.Recordset, Grid As ListView, CountFields As Integer)
On Error Resume Next

Call OpenTable(StrSql, rs)
Grid.ListItems.Clear
Do While Not rs.EOF
   Set List = Grid.ListItems.Add(, , rs.Fields(0))
   For i = 1 To CountFields
      List.SubItems(i) = rs.Fields(i)
   Next i
   rs.MoveNext
Loop
End Sub

 
 

'Public Sub SetFormCenter(Frm As Form)
'Frm.Move (frmUtama.ScaleWidth \ 2) - (Frm.Width \ 2), (frmUtama.ScaleHeight / 2) - (Frm.Height / 2)
'End Sub


  Public Sub Loadkd_plgnToCombo(StrSql As String, rs As ADODB.Recordset, Combo As ComboBox)
Call OpenTable(StrSql, rs)
Combo.Clear
Do While Not rs.EOF
   Combo.AddItem rs.Fields(0)
   rs.MoveNext
Loop
End Sub



Public Sub Loadkd_supToCombo(StrSql As String, rs As ADODB.Recordset, Combo As ComboBox)
Call OpenTable(StrSql, rs)
Combo.Clear
Do While Not rs.EOF
   Combo.AddItem rs.Fields(0)
   rs.MoveNext
Loop
End Sub

Public Sub LoadpembeliToCombo(StrSql As String, rs As ADODB.Recordset, Combo As ComboBox)
Call OpenTable(StrSql, rs)
Combo.Clear
Do While Not rs.EOF
   Combo.AddItem rs.Fields(0)
   rs.MoveNext
Loop
End Sub

Public Sub Loadkd_brgToCombo(StrSql As String, rs As ADODB.Recordset, Combo As ComboBox)
Call OpenTable(StrSql, rs)
Combo.Clear
Do While Not rs.EOF
   Combo.AddItem rs.Fields(0)
   rs.MoveNext
Loop
End Sub

Public Sub Kd_distributor_Click(StrSql As String, rs As ADODB.Recordset, Combo As ComboBox)
Call OpenTable(StrSql, rs)
Combo.Clear
Do While Not rs.EOF
   Combo.AddItem rs.Fields(0)
   rs.MoveNext
Loop
End Sub

Public Sub Loadkd_mobilToCombo(StrSql As String, rs As ADODB.Recordset, Combo As ComboBox)
Call OpenTable(StrSql, rs)
Combo.Clear
Do While Not rs.EOF
   Combo.AddItem rs.Fields(0)
   rs.MoveNext
Loop
End Sub


Public Sub Loadno_fakToCombo(StrSql As String, rs As ADODB.Recordset, Combo As ComboBox)
Call OpenTable(StrSql, rs)
Combo.Clear
Do While Not rs.EOF
   Combo.AddItem rs.Fields(0)
   rs.MoveNext
Loop
End Sub

Public Sub OpenTable(StrSql As String, rs As ADODB.Recordset)
    Set rs = New ADODB.Recordset
        If rs.State = adStateOpen Then Set rs = Nothing
        rs.Open StrSql, Conn, adOpenDynamic, adLockOptimistic
        
    
End Sub

 
 
 Public Sub PesanSudahAda(Frm As Form)
 MsgBox "Data sudah ada!", vbCritical, "Data Suda Ada"
 End Sub
 Public Sub PesanKosong(Frm As Form)
 MsgBox "Data tidak boleh kosong!", vbCritical, "Data Kosong"
  
 End Sub
 

 Public Sub PesanTdkDitemukan(Frm As Form)
 MsgBox "Data tidak ditemukan!", vbCritical, "Cari Data"
  
 End Sub
 
 Public Sub PesanSimpan(Frm As Form)
 MsgBox "Data sudah disimpan!", vbInformation, "Simpan Data"
 End Sub
  Public Sub PesanUpdate(Frm As Form)
 MsgBox "Data sudah di-update!", vbInformation, "Update Data"
 End Sub
 
 Public Sub PesanHapus(Frm As Form)
 MsgBox "Data sudah terhapus!", vbInformation, "Hapus Data"
 End Sub
 
 Public Sub IsiDataText1()
     Isitext = "AaBbCcDdEeFfGgHhIiJjKkLlMmNnOoPpQqRrSsTtUuVvWwXxYyZz.,"
 End Sub
Public Sub IsiDataText2()
     Isitext = "0123456789"
End Sub
Public Sub IsiDataText3()
     Isitext = "()-0123456789"
End Sub
 
 
        
   







