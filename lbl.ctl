VERSION 5.00
Begin VB.UserControl ThreeDLabel 
   ClientHeight    =   240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1440
   ScaleHeight     =   16
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   96
   Begin VB.Label Reference 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Reference"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "ThreeDLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Const m_def_Caption = ""
Const m_def_Alignment = 0
Const m_def_Offset = 1
Dim m_Caption As String
Dim m_Alignment As Integer
Dim m_Offset As Integer

Public Enum Aligner
    [Left Justified] = 0
    [Right Justified] = 1
    [Center] = 2
End Enum

Public Enum Borderer
    [None] = 0
    [Fixed Single] = 1
End Enum
    
Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event Resize()

Public Property Get Alignment() As Aligner
    Alignment = Label1(0).Alignment
End Property
'
Public Property Let Alignment(ByVal New_Alignment As Aligner)
    Label1(0).Alignment() = New_Alignment
    UserControl_Resize
    PropertyChanged "Alignment"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get Offset() As Integer
    Offset = m_Offset
End Property

Public Property Let Offset(New_Offset As Integer)
    m_Offset = New_Offset
    UserControl_Resize
    PropertyChanged "Offset"
End Property

Public Property Get BorderStyle() As Borderer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Borderer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Public Property Get Caption() As String
    Caption = Label1(0).Caption
End Property
'
Public Property Let Caption(ByVal New_Caption As String)
    Label1(0).Caption() = New_Caption
    UserControl_Resize
    PropertyChanged "Caption"
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = Reference.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Reference.Font = New_Font
    UserControl_Resize
    PropertyChanged "Font"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = Reference.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Reference.ForeColor() = New_ForeColor
    UserControl_Resize
    PropertyChanged "ForeColor"
End Property

Public Property Get ShadowColor() As OLE_COLOR
    ShadowColor = Reference.BackColor
End Property

Public Property Let ShadowColor(New_ShadowColor As OLE_COLOR)
    Reference.BackColor = New_ShadowColor
    UserControl_Resize
    PropertyChanged "ShadowColor"
End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hWnd = UserControl.hWnd
End Property

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = UserControl.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set UserControl.Picture = New_Picture
    PropertyChanged "Picture"
End Property

Private Sub UserControl_Resize()
    RaiseEvent Resize
    Label1(1).Caption = Label1(0).Caption
    Set Label1(0).Font = Reference.Font
    Set Label1(1).Font = Reference.Font
    Label1(0).FontBold = Reference.FontBold
    Label1(1).FontBold = Reference.FontBold
    Label1(0).FontItalic = Reference.FontItalic
    Label1(1).FontItalic = Reference.FontItalic
    Label1(0).FontUnderline = Reference.FontUnderline
    Label1(1).FontUnderline = Reference.FontUnderline
    Label1(0).FontStrikethru = Reference.FontStrikethru
    Label1(1).FontStrikethru = Reference.FontStrikethru
    Label1(0).FontName = Reference.FontName
    Label1(1).FontName = Reference.FontName
    Label1(0).Top = 0
    Label1(0).Left = 0
    Label1(0).Width = UserControl.ScaleWidth
    Label1(1).Width = UserControl.ScaleWidth
    Label1(1).Top = m_Offset
    Label1(1).Left = m_Offset
    Label1(0).ForeColor = Reference.ForeColor
    Label1(1).ForeColor = Reference.BackColor
    Label1(1).Alignment = Label1(0).Alignment
    Label1(0).Height = UserControl.ScaleHeight
    Label1(1).Height = UserControl.ScaleHeight
End Sub

Private Sub UserControl_InitProperties()
    m_Caption = m_def_Caption
    Label1(0).Caption = Replace(Ambient.DisplayName, "ThreeDLabel", "Label")
    m_Alignment = m_def_Alignment
    m_Offset = m_def_Offset
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Label1(0).Alignment = PropBag.ReadProperty("Alignment", 0)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    Label1(0).Caption = PropBag.ReadProperty("Caption", "Label1")
    Set Reference.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Reference.ForeColor = PropBag.ReadProperty("ForeColor", &H404040)
    Reference.BackColor = PropBag.ReadProperty("ShadowColor", &HE0E0E0)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    m_Offset = PropBag.ReadProperty("Offset", m_def_Offset)
    m_Caption = PropBag.ReadProperty("Caption", "3D Label")
    m_Alignment = PropBag.ReadProperty("Alignment", m_def_Alignment)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Alignment", Label1(0).Alignment, 0)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("Caption", Label1(0).Caption, "Label1")
    Call PropBag.WriteProperty("Font", Reference.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", Reference.ForeColor, &H404040)
    Call PropBag.WriteProperty("ShadowColor", Reference.BackColor, &HE0E0E0)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("Offset", m_Offset, m_def_Offset)
    Call PropBag.WriteProperty("Caption", Label1(0).Caption, "3D Label")
    Call PropBag.WriteProperty("Alignment", m_Alignment, m_def_Alignment)
End Sub
