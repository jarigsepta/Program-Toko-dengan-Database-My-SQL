VERSION 5.00
Object = "{EAFDAFBF-1D88-41DD-B117-60ECBC4B8441}#1.0#0"; "vkUserControlsXP.ocx"
Begin VB.Form StatusSSH 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Status Berlangganan SSH"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   9930
   StartUpPosition =   1  'CenterOwner
   Begin vkUserContolsXP.vkFrame vkFrame2 
      Height          =   4455
      Left            =   3960
      TabIndex        =   2
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7858
      Caption         =   "Akun SSH"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleColor1     =   16576
      BorderColor     =   49152
      Begin vkUserContolsXP.vkCommand vkdelAccount 
         Height          =   615
         Left            =   4440
         TabIndex        =   19
         Top             =   3720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1085
         Caption         =   "Hapus Akun"
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
      Begin vkUserContolsXP.vkCommand vkgPassword 
         Height          =   615
         Left            =   2880
         TabIndex        =   18
         Top             =   3720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1085
         Caption         =   "Ganti Password"
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
      Begin vkUserContolsXP.vkLabel vkLabel1 
         Height          =   375
         Left            =   4110
         TabIndex        =   15
         Top             =   1860
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   ">"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vkUserContolsXP.vkTextBox vkED 
         Height          =   495
         Left            =   4440
         TabIndex        =   14
         Top             =   1800
         Width           =   1335
         _ExtentX        =   2355
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
         Locked          =   -1  'True
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkTextBox vkAD 
         Height          =   495
         Left            =   2760
         TabIndex        =   13
         Top             =   1800
         Width           =   1215
         _ExtentX        =   2143
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
         Locked          =   -1  'True
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkTextBox vkU 
         Height          =   495
         Left            =   1320
         TabIndex        =   12
         Top             =   2520
         Width           =   2655
         _ExtentX        =   4683
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
         Locked          =   -1  'True
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkLabel vkLabel7 
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2640
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Username :"
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
      Begin vkUserContolsXP.vkLabel vkLabel4 
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Host :"
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
      Begin vkUserContolsXP.vkLabel vkLabel5 
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Port :"
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
      Begin vkUserContolsXP.vkLabel vkLabel6 
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Masa berlaku (hari) :"
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
      Begin vkUserContolsXP.vkLabel vkLabel8 
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   3240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Password :"
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
      Begin vkUserContolsXP.vkTextBox vkHost 
         Height          =   495
         Left            =   1080
         TabIndex        =   6
         Top             =   525
         Width           =   3255
         _ExtentX        =   5741
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
         Locked          =   -1  'True
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkTextBox vkPort 
         Height          =   495
         Left            =   1080
         TabIndex        =   5
         Top             =   1125
         Width           =   1695
         _ExtentX        =   2990
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
         Locked          =   -1  'True
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkTextBox vkD 
         Height          =   495
         Left            =   2160
         TabIndex        =   4
         Top             =   1800
         Width           =   495
         _ExtentX        =   873
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
         Locked          =   -1  'True
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkTextBox vkP 
         Height          =   495
         Left            =   1320
         TabIndex        =   3
         Top             =   3120
         Width           =   2655
         _ExtentX        =   4683
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
         LegendForeColor =   16776960
      End
   End
   Begin vkUserContolsXP.vkFrame vkFrame1 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   7858
      Caption         =   "USERNAME - SSH Server"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleColor1     =   16576
      BorderColor     =   49152
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   3720
         Visible         =   0   'False
         Width           =   2055
      End
      Begin vkUserContolsXP.vkTextBox vkFinDuSSH 
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   3240
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
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
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         LegendForeColor =   16750899
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   3495
      End
   End
End
Attribute VB_Name = "StatusSSH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim servercode As String

Private Sub Combo1_Click()
seeAllActiveSSH
bindAllFromSSH
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = AutoMatchCBBox(Combo1, KeyAscii)
End Sub

Private Sub Form_Activate()
Combo1.SetFocus
End Sub

Private Sub Form_Load()
getNSSH

seeActiveSSH
comboActiveSSH

End Sub

Private Sub seeAllActiveSSH()
sql = "select ssh_item_active.host,ssh_item_active.port,user_ssh,pass_ssh,active_date,exp_date,days_duration from ssh_item_active inner join ssh_item_menu on ssh_item_active.item_id = ssh_item_menu.item_id where server_cuser = '" & Combo1.Text & "' "
Set rs = bridge.Execute(sql)
End Sub

Public Sub getPwd()
sql = "select pass_ssh from ssh_item_active where server_cuser = '" & Combo1.Text & "' "
Set rs = bridge.Execute(sql)
vkP.Text = rs.Fields("pass_ssh").Value
vkP.Refresh
End Sub

Public Sub delAcc()
sql = "delete from ssh_item_active where server_cuser = '" & Combo1.Text & "' "
Set rs = bridge.Execute(sql)
End Sub

Private Sub bindAllFromSSH()
vkHost.Text = rs.Fields("host").Value
vkPort.Text = rs.Fields("port").Value
vkU.Text = rs.Fields("user_ssh").Value
vkP.Text = rs.Fields("pass_ssh").Value
vkAD.Text = rs.Fields("active_date").Value
vkED.Text = rs.Fields("exp_date").Value
vkD.Text = rs.Fields("days_duration").Value
End Sub

Private Sub seeActiveSSH()
'sql = "select ssh_item_menu.ssh_server from ssh_item_menu inner join ssh_item_active on ssh_item_active.item_id = ssh_item_menu.item_id where ssh_item_active.uid = '" & gSID & "' "
sql = "select server_cuser from ssh_item_active where uid = '" & gSID & "' "
Set rs = bridge.Execute(sql)
End Sub

Private Sub comboActiveSSH()
With rs
If .RecordCount > 0 Then
  Do While Not rs.EOF
    Combo1.AddItem rs.Fields("server_cuser").Value
    rs.MoveNext
  Loop
Set rs = Nothing
End If
End With

End Sub

Private Sub Text1_Change()
'CariIntel Text1, "ssh_item_active", "server_cuser"
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
'TxtAutoComplete Text1, KeyAscii
End Sub

Private Sub vkdelAccount_Click()

If MsgBox("Yakin ingin menghapus akun ini?", vbYesNo) = vbYes Then
delAcc
Combo1.Text = ""
vkHost.Text = ""
vkPort.Text = ""
vkU.Text = ""
vkP.Text = ""
vkAD.Text = ""
vkED.Text = ""
vkD.Text = ""
Combo1.Clear
seeActiveSSH
comboActiveSSH
End If

End Sub

Private Sub vkFinDuSSH_KeyDown(KeyCode As Integer, Shift As Integer)
'CheKey KeyCode, Text1
End Sub

Private Sub vkgPassword_Click()
If Combo1.Text = "" Then
MsgBox "Pilih akun yang akan diganti passwordnya"
Else
 vkP.BorderColor = &HC000&
 vkP.SetFocus
    If MsgBox("Lanjutkan", vbYesNo) <> vbNo Then
        ChangePass.Show
    End If
 getPwd
 vkP.BorderColor = &HFF8080
End If
End Sub
