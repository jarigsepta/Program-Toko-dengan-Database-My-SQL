VERSION 5.00
Object = "{EAFDAFBF-1D88-41DD-B117-60ECBC4B8441}#1.0#0"; "vkUserControlsXP.ocx"
Begin VB.Form Login 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4140
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   4140
   StartUpPosition =   2  'CenterScreen
   Begin vkUserContolsXP.vkCommand vkCommand2 
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "Exit"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin vkUserContolsXP.vkCommand cmdLogin 
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "Login"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin vkUserContolsXP.vkLabel vkLabel2 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1150
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BackColor       =   16777215
      BackStyle       =   0
      Caption         =   "Password"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin vkUserContolsXP.vkTextBox vkTextBox1 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      BeginProperty LegendFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      LegendForeColor =   16750899
   End
   Begin vkUserContolsXP.vkLabel vkLabel1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   555
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BackColor       =   16777215
      BackStyle       =   0
      Caption         =   "User/Email"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin vkUserContolsXP.vkTextBox vkTextBox2 
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   1080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      BeginProperty LegendFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      PassWordChar    =   "•"
      LegendForeColor =   16750899
   End
   Begin vkUserContolsXP.vkFrame vkFrame1 
      Height          =   2535
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4471
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin InformationSystem.Hyperlink Hyperlink1 
         Height          =   255
         Left            =   1800
         TabIndex        =   7
         Top             =   2160
         Width           =   2055
         _extentx        =   3625
         _extenty        =   450
         text            =   "Belum punya akun? Daftar!"
         font            =   "Login.frx":0000
         hyperlinkaddress=   "vb"
         backcolor       =   -2147483626
      End
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public gN As String
'Public gX_UID

Dim Enc As CMD5

Private Sub cmdLogin_Click()
Set Enc = New CMD5

If vkTextBox1.Text = "" Or vkTextBox2.Text = "" Then
MsgBox "Silahkan diperiksa ", vbExclamation, "Eror"
Else
sql = "select email,pwd from customer where email = ('" & vkTextBox1.Text & "') AND pwd = ('" & Enc.MD5(vkTextBox2.Text) & "')"
Set rs = bridge.Execute(sql)
'//////////
If (rs.RecordCount) Then
'MsgBox "Successful!", vbInformation
Me.Hide
Main.Show
'get nama user dari database
getName
'getX_UID
Main.Caption = "Selamat Datang, " & gN
Else
MsgBox "Failed!", vbCritical, "Login failed"
End If
'//////////
End If
End Sub

Private Sub Form_Load()
Call CreateConn
End Sub

Private Sub getName()
sql = "select name from customer where email = ('" & vkTextBox1.Text & "')"
Set rs = bridge.Execute(sql)
gN = rs.GetString
End Sub

'Private Sub getX_UID()
'sql = "select uid from customer where email = ('" & vkTextBox1.Text & "')"
'Set rs = bridge.Execute(sql)
'gX_UID = rs.GetString
'End Sub

Private Sub vkCommand2_Click()
Unload Me
End Sub

