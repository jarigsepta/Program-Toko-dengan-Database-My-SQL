VERSION 5.00
Object = "{EAFDAFBF-1D88-41DD-B117-60ECBC4B8441}#1.0#0"; "vkUserControlsXP.ocx"
Begin VB.Form ChangePass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ganti Password"
   ClientHeight    =   915
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5220
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   915
   ScaleWidth      =   5220
   StartUpPosition =   1  'CenterOwner
   Begin vkUserContolsXP.vkCommand vkConfirmChange 
      Height          =   615
      Left            =   4320
      TabIndex        =   2
      Top             =   170
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1085
      Caption         =   "OK"
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
   Begin vkUserContolsXP.vkTextBox vkNewPass 
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      BeginProperty LegendFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LegendForeColor =   16750899
   End
   Begin vkUserContolsXP.vkLabel vkLabel1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BackColor       =   16777215
      BackStyle       =   0
      Caption         =   "Password baru :"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "ChangePass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub vkConfirmChange_Click()
If vkNewPass.Text = "" Then
MsgBox "Ketikkan password baru", vbExclamation
Else
sql = "update ssh_item_active set pass_ssh = '" & vkNewPass.Text & "' where server_cuser = '" & StatusSSH.Combo1.Text & "' "
Set rs = bridge.Execute(sql)
MsgBox ("Password telah diganti!")
StatusSSH.getPwd
Unload Me
End If
End Sub

