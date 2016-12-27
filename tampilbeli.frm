VERSION 5.00
Begin VB.Form Tampilkan 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ESC = Tutup ** Enter = Cetak"
   ClientHeight    =   5730
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   4320
End
Attribute VB_Name = "Tampilkan"
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


Private Sub Form_KeyPress(Keyascii As Integer)
If Keyascii = 27 Then
    Unload Me
ElseIf Keyascii = 13 Then
    pesan = MsgBox("Printer sudah siap", vbYesNo)
    If pesan = vbYes Then
        Call Cetak
    Else
        Unload Me
    End If
End If
End Sub



