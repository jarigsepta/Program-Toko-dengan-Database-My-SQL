VERSION 5.00
Object = "{EAFDAFBF-1D88-41DD-B117-60ECBC4B8441}#1.0#0"; "vkUserControlsXP.ocx"
Begin VB.Form IsiBalance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Isi Ulang Saldo"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5580
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   5580
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox vkPm 
      Height          =   645
      Left            =   1800
      TabIndex        =   5
      Top             =   840
      Width           =   2055
   End
   Begin vkUserContolsXP.vkCommand vkRecharge 
      Height          =   855
      Left            =   3960
      TabIndex        =   2
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1508
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
   Begin vkUserContolsXP.vkLabel vkLabel1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BackColor       =   16777215
      BackStyle       =   0
      Caption         =   "Nominal"
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
   Begin vkUserContolsXP.vkTextBox vkNominal 
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   240
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
      LegendForeColor =   16750899
   End
   Begin vkUserContolsXP.vkCommand vkCommand2 
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
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
   Begin vkUserContolsXP.vkLabel vkLabel2 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BackColor       =   16777215
      BackStyle       =   0
      Caption         =   "Metode Pembayaran"
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
Attribute VB_Name = "IsiBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim valuePm As String
Dim originalBalance As String
Dim addBalance As String

Private Sub Form_Load()
With vkPm
.AddItem "ATM"
.AddItem "Transfer Bank"
.AddItem "Pulsa"
.AddItem "Indomaret"
End With

getNSSH
getManyBalance
End Sub

Private Sub vkCommand2_Click()
Unload Me
End Sub

Private Sub vkNominal_Change()
vkNominal.Text = Format(vkNominal.Text, "###,###,###")
End Sub

Private Sub vkPm_Click()
valuePm = vkPm.Text
End Sub

Private Sub vkRecharge_Click()
If vkNominal.Text = "" Or vkPm.SelCount = 0 Then
MsgBox ("Silahkan masukkan nominal dan metode pembayaran anda")
Else

If MsgBox("Lanjutkan isi ulang", vbYesNo) <> vbNo Then
    Sql = "update customer_wallet set balance = '" & Val(Replace(vkNominal.Text, ".", "")) + originalBalance & "', payment_method = '" & valuePm & "' where uid = '" & gSID & "' "
    'MsgBox (Sql)
    Set RS = bridge.Execute(Sql)
    MsgBox "Isi ulang berhasil", vbInformation
    Unload Me
End If

End If
End Sub

Private Sub getManyBalance()
Sql = "select balance from customer_wallet where uid = '" & gSID & "' "
Set RS = bridge.Execute(Sql)
originalBalance = RS.GetString
End Sub
