VERSION 5.00
Begin VB.Form frmlogin 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login System"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4140
   Icon            =   "frmlogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmlogin.frx":000C
   ScaleHeight     =   2400
   ScaleWidth      =   4140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   150
      Left            =   120
      Top             =   1800
   End
   Begin Project1.Button btncancel 
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   1680
      Width           =   855
      _extentx        =   1508
      _extenty        =   661
      font            =   "frmlogin.frx":22C2B
      caption         =   "Cancel"
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
   Begin Project1.Button btnlogin 
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   1680
      Width           =   855
      _extentx        =   1508
      _extenty        =   661
      font            =   "frmlogin.frx":22C53
      caption         =   "Login"
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
   Begin VB.TextBox Text2 
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
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strTemp, LenTemp, n, enTemp
Dim RS As New ADODB.Recordset
Sub status()
Set RS = New ADODB.Recordset
RS.CursorLocation = adUseClient
End Sub
Private Sub btncancel_Click()
Unload Me
Unload menuutama
End Sub

Private Sub btnlogin_Click()
    status
    Dim p
    p = "select * from [user] where Username='" & Trim(Text1.Text) & "' And Password='" & Trim(Text2.Text) & "'"
    RS.Open p, db
    If RS.EOF Then
        MsgBox "Username atau Password salah !", vbCritical + vbOKOnly, "Peringatan !"
        Text1.Text = ""
        Text2.Text = ""
        Text1.SetFocus
        Else
        MsgBox "Apakah Anda sudah siap ?", vbQuestion + vbOKCancel
        If vbYes Then
        menuutama.Show
        menuutama.StatusBar1.Panels(2) = RS.Fields(2)
        menuutama.StatusBar1.Panels(3) = RS.Fields(4)
        If menuutama.StatusBar1.Panels(3) = "KASIR" Then
            menuutama.mnuopt.Visible = False
            menuutama.mnulap.Visible = False
        End If
        Unload Me
        End If
    End If
End Sub

Private Sub Form_Load()
Call koneksi
strTemp = Me.Caption
    n = 1
Text2.Enabled = False
End Sub

Private Sub Text1_Change()
Dim posisi As Integer
  posisi = Text1.SelStart
  Text1.Text = AwalKataKapital(Text1.Text)
  Text1.SelStart = posisi
End Sub

Private Sub Text1_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    If Text1.Text = "" Then
        MsgBox "Isikan Username dahulu !", vbCritical, "Kesalahan !"
        Text1.SetFocus
        Else
        Text2.Enabled = True
        Text2.SetFocus
        End If
End If
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
