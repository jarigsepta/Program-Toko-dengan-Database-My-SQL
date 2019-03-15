VERSION 5.00
Object = "{EAFDAFBF-1D88-41DD-B117-60ECBC4B8441}#1.0#0"; "vkUserControlsXP.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form PaketSSH 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Paket SSH"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   9135
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   23
      Top             =   5280
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin vkUserContolsXP.vkLabel vkLabel10 
      Height          =   255
      Left            =   100
      TabIndex        =   22
      Top             =   1970
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      BackColor       =   16777215
      BackStyle       =   0
      Caption         =   "Stock :"
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
   Begin vkUserContolsXP.vkTextBox vkBalance 
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
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
         Name            =   "Myriad Hebrew"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      LegendAlignmentX=   0
      LegendForeColor =   16750899
   End
   Begin vkUserContolsXP.vkLabel vkLabel1 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   180
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      BackColor       =   16777215
      BackStyle       =   0
      Caption         =   "Saldo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Myriad Hebrew"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
   End
   Begin vkUserContolsXP.vkFrame vkFrame1 
      Height          =   4575
      Left            =   30
      TabIndex        =   0
      Top             =   720
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   8070
      Caption         =   "SSH"
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
      Begin vkUserContolsXP.vkTextBox vkStock 
         Height          =   495
         Left            =   120
         TabIndex        =   21
         Top             =   1485
         Width           =   855
         _ExtentX        =   1508
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
            Name            =   "Myriad Hebrew"
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
      Begin vkUserContolsXP.vkLabel vkLabel3 
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2145
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Deskripsi :"
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
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   340
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Pilih paket SSH :"
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
      Begin vkUserContolsXP.vkTextBox vkDesc 
         Height          =   2055
         Left            =   120
         TabIndex        =   5
         Top             =   2400
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   3625
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
            Name            =   "Myriad Hebrew"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2
         LegendForeColor =   16750899
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   3375
      End
   End
   Begin vkUserContolsXP.vkFrame vkFrame2 
      Height          =   4575
      Left            =   3720
      TabIndex        =   1
      Top             =   720
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   8070
      Caption         =   "Detail SSH"
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
      Begin vkUserContolsXP.vkCommand vkBuy 
         Height          =   495
         Left            =   3840
         TabIndex        =   20
         Top             =   3840
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         Caption         =   "Beli"
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
      Begin vkUserContolsXP.vkTextBox vkTotal 
         Height          =   495
         Left            =   1080
         TabIndex        =   19
         Top             =   3000
         Width           =   2895
         _ExtentX        =   5106
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
            Name            =   "Myriad Hebrew"
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
      Begin vkUserContolsXP.vkLabel vkLabel9 
         Height          =   255
         Left            =   3050
         TabIndex        =   18
         Top             =   2530
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "X"
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
      Begin vkUserContolsXP.vkTextBox vkMany 
         Height          =   495
         Left            =   3360
         TabIndex        =   17
         Top             =   2400
         Width           =   615
         _ExtentX        =   1085
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
            Name            =   "Myriad Hebrew"
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
      Begin vkUserContolsXP.vkTextBox vkPrice 
         Height          =   495
         Left            =   1080
         TabIndex        =   16
         Top             =   2400
         Width           =   1815
         _ExtentX        =   3201
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
            Name            =   "Myriad Hebrew"
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
      Begin MSComctlLib.Slider Slider1 
         Height          =   375
         Left            =   2040
         TabIndex        =   15
         Top             =   1785
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         Min             =   1
         Max             =   7
         SelStart        =   1
         Value           =   1
      End
      Begin vkUserContolsXP.vkTextBox vkPort 
         Height          =   495
         Left            =   1080
         TabIndex        =   14
         Top             =   1125
         Width           =   1455
         _ExtentX        =   2566
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
            Name            =   "Myriad Hebrew"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         LegendForeColor =   12582912
      End
      Begin vkUserContolsXP.vkTextBox vkHost 
         Height          =   495
         Left            =   1080
         TabIndex        =   13
         Top             =   525
         Width           =   3015
         _ExtentX        =   5318
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
            Name            =   "Myriad Hebrew"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         LegendForeColor =   16711680
      End
      Begin vkUserContolsXP.vkLabel vkLabel8 
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   3090
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Jumlah :"
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
      Begin vkUserContolsXP.vkLabel vkLabel7 
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2470
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Harga :"
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
         TabIndex        =   10
         Top             =   1800
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Masa berlaku (hari):"
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
         Top             =   1230
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
      Begin vkUserContolsXP.vkLabel vkLabel4 
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   610
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
   End
End
Attribute VB_Name = "PaketSSH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim kx As String

Private Sub Combo1_Click()
On Error Resume Next
getAllFromSSH 'get all ssh data
putAllFromSSH 'put all ssh data
vkTotal.Text = Val(vkPrice.Text * vkMany.Text) 'rebind data
vkPrice.Text = Format(vkPrice.Text, "###,###,###")
vkTotal.Text = Format(vkTotal.Text, "###,###,###")

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = AutoMatchCBBox(Combo1, KeyAscii)
End Sub

Private Sub Form_Activate()
Combo1.SetFocus
End Sub

Private Sub Form_Load()

getSSH 'get ssh Name
putSSH 'put ssh Name
getNSSH 'extinct call uid from customer in module getID
getBalance 'get balance user

StatusBar1.Panels(1).Width = 2500

vkBalance.Text = Format(vkBalance.Text, "###,###,###")
vkTotal.Text = Format(vkTotal.Text, "###,###,###")
vkPrice.Text = Format(vkPrice.Text, "###,###,###")

End Sub

Private Sub Slider1_Change()
On Error Resume Next
vkPrice.Text = Format(vkPrice.Text, "###,###,###")
vkMany.Text = Slider1.Value
vkTotal.Text = Val(vkPrice.Text * vkMany.Text)
vkTotal.Text = Format(vkTotal.Text, "###,###,###")
End Sub

Private Sub getAllFromSSH()
sql = "select host,port,price,stock,description from ssh_item_menu where ssh_server = '" & Combo1.Text & "' "
Set rs = bridge.Execute(sql)
End Sub

Private Sub putAllFromSSH()
vkHost.Text = rs.Fields("host").Value
vkPort.Text = rs.Fields("port").Value
vkPrice.Text = rs.Fields("price").Value
vkStock.Text = rs.Fields("stock").Value
vkDesc.Text = rs.Fields("description").Value

Set rs = Nothing
End Sub

Public Sub getBalance()
On Error Resume Next
sql = "select customer_wallet.balance from customer_wallet left outer join customer on customer_wallet.uid = customer.uid where customer_wallet.uid = '" & gSID & "' "
Set rs = bridge.Execute(sql)
vkBalance.Text = rs.GetString
vkBalance.Text = Format(vkBalance.Text, "###,###,###")
End Sub

Public Sub getStock()
On Error Resume Next
sql = "select stock from ssh_item_menu where ssh_server = '" & Combo1.Text & "' "
Set rs = bridge.Execute(sql)
vkStock.Text = rs.GetString
End Sub

Private Sub getSSH()
sql = "select ssh_server from ssh_item_menu"
Set rs = bridge.Execute(sql)
End Sub

Private Sub putSSH()
With rs
.MoveFirst
If .RecordCount > 0 Then
  Do While Not rs.EOF
    Combo1.AddItem rs.Fields("ssh_server").Value
    rs.MoveNext
  Loop

Set rs = Nothing
End If
End With

End Sub

Private Sub vkBuy_Click()

If Val(Replace(vkBalance.Text, ".", "")) < Val(Replace(vkTotal.Text, ".", "")) Then
  StatusBar1.Panels(1).Text = "Saldo tidak cukup!"
ElseIf vkStock.Text = 0 Then
    StatusBar1.Panels(1).Text = "Stock habis!"
ElseIf vkMany.Text = "" Then
  StatusBar1.Panels(1).Text = "Ingin beli berapa?"
Else
  StatusBar1.Panels(1).Text = "Pembelian diproses.."
  Auth.Show
End If

End Sub
