VERSION 5.00
Object = "{EAFDAFBF-1D88-41DD-B117-60ECBC4B8441}#1.0#0"; "vkUserControlsXP.ocx"
Begin VB.Form Auth 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create SSH Account"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4665
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   4665
   StartUpPosition =   1  'CenterOwner
   Begin vkUserContolsXP.vkCommand vCreateAcc 
      Height          =   495
      Left            =   3480
      TabIndex        =   4
      Top             =   160
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
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
   Begin vkUserContolsXP.vkLabel vkLabel2 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   870
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
   Begin vkUserContolsXP.vkLabel vkLabel1 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   270
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BackColor       =   16777215
      BackStyle       =   0
      Caption         =   "Username"
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
   Begin vkUserContolsXP.vkTextBox vkPassSSH 
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   840
      Width           =   2175
      _ExtentX        =   3836
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
      LegendForeColor =   16750899
   End
   Begin vkUserContolsXP.vkTextBox vkUserSSH 
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   2175
      _ExtentX        =   3836
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
      LegendForeColor =   16750899
   End
   Begin vkUserContolsXP.vkCommand vkCancelTX 
      Height          =   495
      Left            =   3480
      TabIndex        =   5
      Top             =   810
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      Caption         =   "Cancel"
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
End
Attribute VB_Name = "Auth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim INsqlActiveSSH As String
Dim INsqlStockServerSSH As String
Dim INsqlTransaction As String
Dim INsqlBalance As String

Dim dayActiveSSH As String
Dim dayAddSSH As String
Dim dayExpireSSH As String

Dim dayTimeTransaction As String

Dim vy As Integer
Dim cUserSSH

Private Sub Form_Load()
getNSSH
getActiveIDSSH

dayActiveSSH = Format(Now, "yyyy/mm/dd")
dayAddSSH = DateAdd("d", Val(PaketSSH.vkMany.Text), dayActiveSSH)
dayExpireSSH = Format(dayAddSSH, "yyyy/mm/dd")
dayTimeTransaction = Format(Time, "hh:mm:ss")

getVY
End Sub

Private Sub vCreateAcc_Click()
Update_X
Insert_XActiveSSH
Insert_XTransaction
PaketSSH.getBalance
PaketSSH.getStock

PaketSSH.StatusBar1.Panels(1).Text = "Pembelian sukses"
Unload Me
End Sub

Private Sub vkCancelTX_Click()
If MsgBox("Batalkan pembelian?", vbOKCancel) = vbOK Then
Unload Me
PaketSSH.StatusBar1.Panels(1).Text = "Pembelian dibatalkan"
End If
End Sub

Private Sub Update_X()
INsqlStockServerSSH = "update ssh_item_menu set stock = '" & Val(PaketSSH.vkStock.Text) - 1 & "' where ssh_server = '" & PaketSSH.Combo1.Text & "' "
INsqlBalance = "update customer_wallet set balance = " & Val(Replace(PaketSSH.vkBalance.Text, ".", "")) - Val(Replace(PaketSSH.vkTotal.Text, ".", "")) & " where uid = '" & gSID & "' "
'MsgBox INsqlStockServerSSH & " " & INsqlBalance
Set rs = bridge.Execute(INsqlStockServerSSH)
Set rs = bridge.Execute(INsqlBalance)
End Sub

Private Sub Insert_XTransaction()
INsqlTransaction = "insert into trx values ('" & gTXID & "','" & Val(Replace(PaketSSH.vkTotal.Text, ".", "")) & "','" & dayActiveSSH & "','" & dayTimeTransaction & "','" & gSID & "','" & Login.gN & "','" & gActiveID & "','" & PaketSSH.Combo1.Text & "','" & dayActiveSSH & "','" & dayExpireSSH & "') "
Set rs = bridge.Execute(INsqlTransaction)
End Sub

Private Sub getVY()
getTXID
vy = Register.PurgeNumericInput(Right(rs.GetString, 2))
gTXID = "TX" & Year(Now) & vy + 1
End Sub

Private Sub Insert_XActiveSSH()
cUserSSH = "[" & vkUserSSH.Text & "] - " & PaketSSH.Combo1.Text
INsqlActiveSSH = "insert into ssh_item_active values ('" & gActiveID & "','" & gSID & "','" & cUserSSH & "','" & PaketSSH.vkHost.Text & "','" & PaketSSH.vkPort.Text & "','" & vkUserSSH.Text & "','" & vkPassSSH.Text & "','" & dayActiveSSH & "','" & dayExpireSSH & "','" & PaketSSH.vkMany.Text & "') "
Set rs = bridge.Execute(INsqlActiveSSH)
End Sub
