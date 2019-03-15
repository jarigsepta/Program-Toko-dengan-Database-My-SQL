VERSION 5.00
Begin VB.UserControl Hyperlink 
   AutoRedraw      =   -1  'True
   ClientHeight    =   285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2205
   ScaleHeight     =   285
   ScaleWidth      =   2205
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   135
      ScaleHeight     =   510
      ScaleWidth      =   915
      TabIndex        =   0
      Top             =   315
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   180
      Top             =   0
   End
End
Attribute VB_Name = "Hyperlink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'api types
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

'constants
Private Const DT_SINGLELINE As Long = &H20
Private Const DT_VCENTER As Long = &H4
Private Const DT_BOTTOM As Long = &H8
Private Const DT_CENTER As Long = &H1
Private Const DT_LEFT As Long = &H0
Private Const DT_RIGHT As Long = &H2
Private Const DT_TOP As Long = &H0
Private Const SRCCOPY As Long = &HCC0020

'enumerations
Public Enum enVAlign
   alTop
   alVCenter
   alBottom
End Enum

Public Enum enHAlign
   alLeft
   alHCenter
   alRight
End Enum

Public Enum enLB
   alwaysUnderline
   hoverUnderline
   neverUnderline
End Enum

Public Enum enBordStyle
   None
   Sunken
   Raised
End Enum

'api calls
Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

'Default Property Values:
Const m_def_ImageLeft = 5
Const m_def_ImageTop = 5
Const m_def_ImageWidth = 32
Const m_def_ImageHeight = 32
Const m_def_BorderStyle = 0
Const m_def_hyperlinkAddress = "http://www.planetsourcecode.com/vb"
Const m_def_activeLinkColor = vbRed
Const m_def_linkBehavior = 1
Const m_def_linkColor = vbBlue
Const m_def_VTextAlign = 1
Const m_def_HTextAlign = 1
Const m_def_Text = "visit planetsourcecode :-)"

'Property Variables:
Dim m_ImageLeft        As Long
Dim m_ImageTop         As Long
Dim m_ImageWidth       As Long
Dim m_ImageHeight      As Long
Dim m_BorderStyle      As enBordStyle
Dim m_hyperlinkAddress As String
Dim m_activeLinkColor  As OLE_COLOR
Dim m_linkBehavior     As enLB
Dim m_linkColor        As OLE_COLOR
Dim m_VTextAlign       As enVAlign
Dim m_HTextAlign       As enHAlign
Dim m_Text             As String

'member variables not attached to property
Dim m_b_mousedown      As Boolean


'================================================
'    SUBS AND FUNCTIONS
'================================================

Private Sub subPaintPicture()

  With Picture1
     StretchBlt hdc, m_ImageLeft, m_ImageTop, _
                m_ImageWidth, m_ImageHeight, _
               .hdc, 0, 0, (.Width \ Screen.TwipsPerPixelX), _
               (.Height \ Screen.TwipsPerPixelY), _
                SRCCOPY
    End With
End Sub
Private Sub subPaintText()
Dim tlen  As Long
Dim lD_VT  As Long
Dim lD_HT  As Long
Dim lDT    As Long

  'set alignment vertically
  If m_VTextAlign = alVCenter Then
     lD_VT = DT_VCENTER
  ElseIf m_VTextAlign = alTop Then
     lD_VT = DT_TOP
  ElseIf m_VTextAlign = alBottom Then
     lD_VT = DT_BOTTOM
  End If
  
  'set alignment horizontally
  If m_HTextAlign = alHCenter Then
     lD_HT = DT_CENTER
  ElseIf m_HTextAlign = alLeft Then
     lD_HT = DT_LEFT
  ElseIf m_HTextAlign = alRight Then
     lD_HT = DT_RIGHT
  End If
  
  'combine the horizontal and vertical aligning
  lDT = (lD_VT Or lD_HT Or DT_SINGLELINE)
  'erase the control
  UserControl.Cls
  tlen = Len(m_Text)
  
  If m_b_mousedown = True Then
    'set the text color [m_activeLinkColor]
    '(mouse is down)
     SetTextColor UserControl.hdc, m_activeLinkColor
  Else
    'set the text color [m_linkColor]
    '(mouse is up)
     SetTextColor UserControl.hdc, m_linkColor
  End If
  
  'draw the text
  DrawText hdc, m_Text, tlen, funcCreateRect, lDT
  
  'if borderstyle is raised then paint
  'the highlights and shadows
  If m_BorderStyle = Raised Then
    With UserControl
       UserControl.Line (0, 0)-(.Width, .Height), _
                      vbWhite, B
       UserControl.Line (-100, -100)-(.Width - 20, _
            .Height - 20), RGB(130, 130, 160), B
    End With
  End If
   
  'paint the image if there is one to paint
  If Not (Picture1.Picture = 0) Then
     subPaintPicture
  End If
End Sub
 
Private Function funcCreateRect() As RECT
  'create rect based upon
  'usercontrols width + height
  With funcCreateRect
    .Left = 0
    .Top = 0
    .Right = (UserControl.Width / Screen.TwipsPerPixelX)
    .Bottom = (UserControl.Height / Screen.TwipsPerPixelY)
  End With
End Function


'================================================
'    CONTROLS CODE
'================================================
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If m_linkBehavior = hoverUnderline Then
     UserControl.FontUnderline = True
     subPaintText
     Timer1 = True
  End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   '
   m_b_mousedown = True
   subPaintText
   'go to the webpage of the texts link
   Register.Show
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   '
   m_b_mousedown = False
   subPaintText
End Sub

Private Sub UserControl_Resize()
   subPaintText
End Sub

Private Sub UserControl_Show()
   subPaintText
End Sub

Private Sub Timer1_Timer()
Dim pt As POINTAPI
  '
  'this timer is running because
  'm_linkBehavior = hoverUnderline
  'and the mouse moved over this control
  '
  'if mouse left this control, turn off
  'the underlining and shut this timer off
  GetCursorPos pt
  If WindowFromPoint(pt.x, pt.y) <> UserControl.hWnd Then
    UserControl.FontUnderline = False
    Timer1.Enabled = False
    subPaintText
  End If

End Sub


'================================================
'    PROPERTIES CODE
'================================================
Public Property Get activeLinkColor() As OLE_COLOR
Attribute activeLinkColor.VB_Description = "Color of controls text when clicked on"
    activeLinkColor = m_activeLinkColor
End Property
Public Property Let activeLinkColor(ByVal New_activeLinkColor As OLE_COLOR)
    m_activeLinkColor = New_activeLinkColor
    PropertyChanged "activeLinkColor"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    subPaintText
End Property
 
Public Property Get BorderStyle() As enBordStyle
    BorderStyle = m_BorderStyle
End Property
Public Property Let BorderStyle(ByVal New_BorderStyle As enBordStyle)
    m_BorderStyle = New_BorderStyle
    
    With UserControl
      If New_BorderStyle = None Then
         .BorderStyle = 0
      ElseIf New_BorderStyle = Sunken Then
         .BorderStyle = 1
      ElseIf New_BorderStyle = Raised Then
         .BorderStyle = 0
      End If
    End With
    
    PropertyChanged "BorderStyle"
    subPaintText
End Property

Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property
Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    subPaintText
End Property
 
Public Property Get HTextAlign() As enHAlign
Attribute HTextAlign.VB_Description = "How the text within the control is aligned horizontally"
    HTextAlign = m_HTextAlign
End Property
Public Property Let HTextAlign(ByVal New_HTextAlign As enHAlign)
    m_HTextAlign = New_HTextAlign
    PropertyChanged "HTextAlign"
    subPaintText
End Property

Public Property Get hyperlinkAddress() As String
Attribute hyperlinkAddress.VB_Description = "The address your default browser navigates to when you click the control (can be a file path as well, i.e C:\\ )"
    hyperlinkAddress = m_hyperlinkAddress
End Property
Public Property Let hyperlinkAddress(ByVal New_hyperlinkAddress As String)
    m_hyperlinkAddress = New_hyperlinkAddress
    PropertyChanged "hyperlinkAddress"
End Property

Public Property Get hdc() As Long
    hdc = UserControl.hdc
End Property
 
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get Image() As Picture
    Set Image = Picture1.Picture
End Property
Public Property Set Image(ByVal New_Image As Picture)
    Set Picture1.Picture = New_Image
    PropertyChanged "Image"
    subPaintText
End Property

Public Property Get ImageLeft() As Long
    ImageLeft = (m_ImageLeft * Screen.TwipsPerPixelX)
End Property
Public Property Let ImageLeft(ByVal New_ImageLeft As Long)
    m_ImageLeft = (New_ImageLeft \ Screen.TwipsPerPixelX)
    PropertyChanged "ImageLeft"
    subPaintText
End Property

Public Property Get ImageTop() As Long
    ImageTop = (m_ImageTop * Screen.TwipsPerPixelY)
End Property
Public Property Let ImageTop(ByVal New_ImageTop As Long)
    m_ImageTop = (New_ImageTop \ Screen.TwipsPerPixelY)
    PropertyChanged "ImageTop"
    subPaintText
End Property

Public Property Get ImageWidth() As Long
    ImageWidth = (m_ImageWidth * Screen.TwipsPerPixelX)
End Property
Public Property Let ImageWidth(ByVal New_ImageWidth As Long)
    m_ImageWidth = (New_ImageWidth \ Screen.TwipsPerPixelX)
    PropertyChanged "ImageWidth"
    subPaintText
End Property

Public Property Get ImageHeight() As Long
    ImageHeight = (m_ImageHeight * Screen.TwipsPerPixelY)
End Property
Public Property Let ImageHeight(ByVal New_ImageHeight As Long)
    m_ImageHeight = (New_ImageHeight \ Screen.TwipsPerPixelY)
    PropertyChanged "ImageHeight"
    subPaintText
End Property

Public Property Get linkBehavior() As enLB
Attribute linkBehavior.VB_Description = "Controls when and if to display the controls text as underlined so it mimicks the behavior of an active link in a webpage"
    linkBehavior = m_linkBehavior
End Property
Public Property Let linkBehavior(ByVal New_linkBehavior As enLB)
    m_linkBehavior = New_linkBehavior
    PropertyChanged "linkBehavior"

    If m_linkBehavior = alwaysUnderline Then
      Font.Underline = True
    Else
      Font.Underline = False
    End If

    subPaintText
End Property

Public Property Get linkColor() As OLE_COLOR
Attribute linkColor.VB_Description = "Color of the text in the default state"
    linkColor = m_linkColor
End Property
Public Property Let linkColor(ByVal New_linkColor As OLE_COLOR)
    m_linkColor = New_linkColor
    PropertyChanged "linkColor"
End Property

Public Property Get Text() As String
Attribute Text.VB_Description = "The text to display in the control. This text also becomes the web address the users default browser navigates to when clicked on"
    Text = m_Text
End Property
Public Property Let Text(ByVal New_Text As String)
    m_Text = New_Text
    PropertyChanged "Text"
    subPaintText
End Property

Public Property Get VTextAlign() As enVAlign
Attribute VTextAlign.VB_Description = "How the text within the control is aligned vertically"
    VTextAlign = m_VTextAlign
End Property
Public Property Let VTextAlign(ByVal New_VTextAlign As enVAlign)
    m_VTextAlign = New_VTextAlign
    PropertyChanged "VTextAlign"
    subPaintText
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Text = m_def_Text
    m_linkColor = m_def_linkColor
    m_linkBehavior = m_def_linkBehavior
    Set UserControl.Font = Ambient.Font
    m_HTextAlign = m_def_HTextAlign
    m_VTextAlign = m_def_VTextAlign
    m_activeLinkColor = m_def_activeLinkColor
    m_hyperlinkAddress = m_def_hyperlinkAddress
    m_BorderStyle = m_def_BorderStyle
    m_ImageLeft = m_def_ImageLeft
    m_ImageTop = m_def_ImageTop
    m_ImageWidth = m_def_ImageWidth
    m_ImageHeight = m_def_ImageHeight
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  With PropBag
     m_Text = .ReadProperty("Text", m_def_Text)
     m_VTextAlign = .ReadProperty("VTextAlign", m_def_VTextAlign)
     m_HTextAlign = .ReadProperty("HTextAlign", m_def_HTextAlign)
     m_linkColor = .ReadProperty("linkColor", m_def_linkColor)
     m_linkBehavior = .ReadProperty("linkBehavior", m_def_linkBehavior)
     Set UserControl.Font = .ReadProperty("Font", Ambient.Font)
     m_activeLinkColor = .ReadProperty("activeLinkColor", m_def_activeLinkColor)
     m_hyperlinkAddress = .ReadProperty("hyperlinkAddress", m_def_hyperlinkAddress)
     UserControl.BackColor = .ReadProperty("BackColor", &H8000000F)
     m_BorderStyle = .ReadProperty("BorderStyle", m_def_BorderStyle)
     Set Picture1.Picture = .ReadProperty("Image", Nothing)
     m_ImageLeft = .ReadProperty("ImageLeft", m_def_ImageLeft)
     m_ImageTop = .ReadProperty("ImageTop", m_def_ImageTop)
     m_ImageWidth = .ReadProperty("ImageWidth", m_def_ImageWidth)
     m_ImageHeight = .ReadProperty("ImageHeight", m_def_ImageHeight)
  End With
 End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  With PropBag
    Call .WriteProperty("Text", m_Text, m_def_Text)
    Call .WriteProperty("VTextAlign", m_VTextAlign, m_def_VTextAlign)
    Call .WriteProperty("HTextAlign", m_HTextAlign, m_def_HTextAlign)
    Call .WriteProperty("linkColor", m_linkColor, m_def_linkColor)
    Call .WriteProperty("linkBehavior", m_linkBehavior, m_def_linkBehavior)
    Call .WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call .WriteProperty("activeLinkColor", m_activeLinkColor, m_def_activeLinkColor)
    Call .WriteProperty("hyperlinkAddress", m_hyperlinkAddress, m_def_hyperlinkAddress)
    Call .WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call .WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call .WriteProperty("Image", Picture1.Picture, Nothing)
    Call .WriteProperty("ImageLeft", m_ImageLeft, m_def_ImageLeft)
    Call .WriteProperty("ImageTop", m_ImageTop, m_def_ImageTop)
    Call .WriteProperty("ImageWidth", m_ImageWidth, m_def_ImageWidth)
    Call .WriteProperty("ImageHeight", m_ImageHeight, m_def_ImageHeight)
  End With
End Sub
  
 

