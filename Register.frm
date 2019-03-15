VERSION 5.00
Object = "{EAFDAFBF-1D88-41DD-B117-60ECBC4B8441}#1.0#0"; "vkUserControlsXP.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Register 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Register"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   8385
   StartUpPosition =   1  'CenterOwner
   Begin vkUserContolsXP.vkCommand vkRegister 
      Height          =   495
      Left            =   6600
      TabIndex        =   21
      Top             =   3430
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      Caption         =   "Register Now"
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
   Begin vkUserContolsXP.vkFrame vkFrame2 
      Height          =   2895
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   5106
      Caption         =   "User Info"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin vkUserContolsXP.vkCheck vkCheck1 
         Height          =   255
         Left            =   1080
         TabIndex        =   11
         Top             =   2480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "WhatsApp?"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   2
      End
      Begin vkUserContolsXP.vkTextBox vkTextBox4 
         Height          =   375
         Left            =   1080
         TabIndex        =   10
         Top             =   1965
         Width           =   1455
         _ExtentX        =   2566
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
      Begin vkUserContolsXP.vkTextBox vkTextBox3 
         Height          =   375
         Left            =   1080
         TabIndex        =   9
         Top             =   1500
         Width           =   1455
         _ExtentX        =   2566
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
      Begin vkUserContolsXP.vkTextBox vkTextBox2 
         Height          =   375
         Left            =   1080
         TabIndex        =   8
         Top             =   1020
         Width           =   1455
         _ExtentX        =   2566
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
      Begin vkUserContolsXP.vkTextBox vkTextBox1 
         Height          =   375
         Left            =   1080
         TabIndex        =   7
         Top             =   555
         Width           =   1455
         _ExtentX        =   2566
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
      Begin vkUserContolsXP.vkLabel vkLabel5 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   2520
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
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
      End
      Begin vkUserContolsXP.vkLabel vkLabel4 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   2040
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "* No HP"
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
      Begin vkUserContolsXP.vkLabel vkLabel3 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "* Password"
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
         TabIndex        =   3
         Top             =   1080
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "* Nama"
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
         TabIndex        =   2
         Top             =   600
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "* Email"
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
   Begin vkUserContolsXP.vkLabel vkLabel6 
      Height          =   375
      Left            =   3720
      TabIndex        =   12
      Top             =   645
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BackColor       =   16777215
      BackStyle       =   0
      Caption         =   "* Alamat"
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
   Begin vkUserContolsXP.vkLabel vkLabel7 
      Height          =   375
      Left            =   3720
      TabIndex        =   13
      Top             =   1125
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BackColor       =   16777215
      BackStyle       =   0
      Caption         =   "* JK"
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
   Begin vkUserContolsXP.vkLabel vkLabel8 
      Height          =   375
      Left            =   3720
      TabIndex        =   14
      Top             =   1605
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BackColor       =   16777215
      BackStyle       =   0
      Caption         =   "* Tgl Lahir"
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
   Begin vkUserContolsXP.vkLabel vkLabel9 
      Height          =   375
      Left            =   3720
      TabIndex        =   15
      Top             =   2085
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BackColor       =   16777215
      BackStyle       =   0
      Caption         =   "Pekerjaan"
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
   Begin vkUserContolsXP.vkTextBox vkTextBox5 
      Height          =   375
      Left            =   4680
      TabIndex        =   16
      Top             =   600
      Width           =   3495
      _ExtentX        =   6165
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
   Begin vkUserContolsXP.vkTextBox vkTextBox8 
      Height          =   375
      Left            =   4680
      TabIndex        =   17
      Top             =   2040
      Width           =   3375
      _ExtentX        =   5953
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
   Begin vkUserContolsXP.vkFrame vkFrame1 
      Height          =   3255
      Left            =   3480
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5741
      Caption         =   "Personal Info"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox CBAgm 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   2430
         Width           =   1935
      End
      Begin vkUserContolsXP.vkCheck vkCheck2 
         Height          =   375
         Left            =   2400
         TabIndex        =   22
         Top             =   2830
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Accept terms and condition"
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
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1200
         TabIndex        =   20
         Top             =   1420
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         Format          =   170459137
         CurrentDate     =   43007
      End
      Begin VB.ComboBox cbgender 
         Height          =   315
         Left            =   1200
         TabIndex        =   19
         Text            =   "Pilih.."
         Top             =   960
         Width           =   2175
      End
      Begin vkUserContolsXP.vkLabel vkLabel10 
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   2460
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Agama"
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
End
Attribute VB_Name = "Register"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sDateString As String
Dim useWa As String

Dim sqlcustomer As String
Dim sqlcustomerpinfo As String
Dim sqlcustomerwallet As String

Dim vz As Integer

Dim Enc As New CMD5

Private Sub DTPicker1_Change()
sDateString = Format(DTPicker1.Value, "yyyy/mm/dd")
End Sub

Private Sub Form_Load()
Call CreateConn

With cbgender
.AddItem "laki-laki"
.AddItem "perempuan"
End With

With CBAgm
.AddItem "Islam"
.AddItem "Kristen"
.AddItem "Katolik"
.AddItem "Hindu"
.AddItem "Buddha"
.AddItem "Other"
End With

vkCheck1.Value = vbUnchecked
useWa = "X"

vkRegister.Enabled = False
End Sub

Private Sub vkCheck1_Change(Value As CheckBoxConstants)
If vkCheck1.Value = vbUnchecked Then
useWa = "X"
Else
useWa = vkTextBox4.Text
End If
End Sub

Private Sub vkCheck2_Change(Value As CheckBoxConstants)
Select Case vkCheck2.Value
Case vbChecked
Agreement.Show
Agreement.vkTextBox1.SetFocus
vkRegister.Enabled = True
Case vbUnchecked
vkRegister.Enabled = False
End Select
End Sub

Private Sub vkRegister_Click()
Set Enc = New CMD5
'///// start dapatkan uid
getUID
vz = PurgeNumericInput(Right(rs.GetString, 3))
gUID = "OH" & Year(Now) & "00" & vz + 1
'///// end dapatkan uid
If vkTextBox1.Text = "" Or vkTextBox2.Text = "" Or vkTextBox3.Text = "" Or vkTextBox4.Text = "" Or vkTextBox5.Text = "" Or cbgender.Text = "Gender" Or DTPicker1.Value = DateValue(Date) Then
MsgBox "Silahkan dilengkapi!", vbExclamation, "Kesalahan"
Else
sqlcustomer = "insert into customer values ('" & gUID & "','" & vkTextBox1.Text & "','" & vkTextBox2.Text & "',('" & Enc.MD5(vkTextBox3.Text) & "'),'" & vkTextBox4.Text & "', '" & useWa & "')"
sqlcustomerpinfo = "insert into customer_pinfo values ('" & gUID & "','" & vkTextBox5.Text & "','" & cbgender.Text & "','" & sDateString & "','" & vkTextBox8.Text & "','" & CBAgm.Text & "')"
sqlcustomerwallet = "insert into customer_wallet values ('" & gUID & "','0','None')"
Set rs = bridge.Execute(sqlcustomer)
Set rs = bridge.Execute(sqlcustomerpinfo)
Set rs = bridge.Execute(sqlcustomerwallet)

MsgBox "Berhasil Mendaftar, silahkan login :)", vbInformation, "Sukses"
Unload Me
Login.Show
End If
End Sub

'////////// Fetch No ID aja
Public Function PurgeNumericInput(ByVal strString As String) As Variant

Dim blnHasDecimal As Boolean, I As Integer
Dim s As String
strString = Trim$(strString)
If Len(strString) = 0 Then
    Exit Function
End If

For I = 1 To Len(strString)
  Select Case Mid$(strString, I, 1)
    Case "0" To "9"
        s = s & Mid$(strString, I, 1)
    Case ".", ","
        If Not blnHasDecimal Then
            blnHasDecimal = True
            s = s & "."
        End If
  End Select
Next I
PurgeNumericInput = Val(s)

End Function
