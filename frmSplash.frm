VERSION 5.00
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   4065
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   7215
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   4065
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4050
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7200
      Begin Project1.VistaForm VistaForm1 
         Height          =   390
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   688
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   0   'False
         ForeColor       =   16777215
         FontItalic      =   0   'False
         FontSize        =   8,25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ShowCloseButton =   0   'False
         ShowMinimiseButton=   0   'False
         ShowMaximiseButton=   0   'False
         Style           =   1
         Transparency    =   -1  'True
      End
      Begin Project1.ProgressBar ProgressBar1 
         Height          =   495
         Left            =   720
         TabIndex        =   5
         Top             =   2280
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   873
         Theme           =   8
         TextStyle       =   3
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "Cheking For Database......"
         TextEffectColor =   16744576
         TextEffect      =   5
      End
      Begin VB.Timer Timer1 
         Interval        =   30
         Left            =   1320
         Top             =   3480
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "System Toko Sepatu"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         TabIndex        =   8
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Universitas Islam Majapahit"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         TabIndex        =   6
         Top             =   960
         Width           =   6495
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Yusman Adi Cahyo (5.14.04.11.0.144)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   4
         Top             =   3600
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sukmana Saputra (5.14.04.11.0.124)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   3
         Top             =   3360
         Width           =   2775
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Di susun oleh :    "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3840
         TabIndex        =   1
         Top             =   2880
         Width           =   2190
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tugas Visual Basic"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   1440
         TabIndex        =   2
         Top             =   360
         Width           =   4080
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub Timer1_Timer()
Static ictr As Integer
If ictr <= 100 Then
ProgressBar1.Value = ictr
ictr = ictr + 1
Else
If ProgressBar1.Max = 100 Then
frmlogin.Show
Unload Me
End If
End If
End Sub
