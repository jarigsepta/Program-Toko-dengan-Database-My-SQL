VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selamat Datang, [User]"
   ClientHeight    =   4680
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   9450
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Main.frx":0000
   ScaleHeight     =   4680
   ScaleWidth      =   9450
   StartUpPosition =   1  'CenterOwner
   Begin VB.Menu menuHosting 
      Caption         =   "SSH"
      Begin VB.Menu menuPktHosting 
         Caption         =   "Beli Paket SSH"
      End
      Begin VB.Menu menuStatHosting 
         Caption         =   "Status Berlangganan SSH"
      End
      Begin VB.Menu mnLogout 
         Caption         =   "Logout"
      End
   End
   Begin VB.Menu menuOpsi 
      Caption         =   "Opsi"
      Begin VB.Menu menuDompet 
         Caption         =   "Dompet"
      End
      Begin VB.Menu menuCtkStruk 
         Caption         =   "Cetak Struk"
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub menuCtkStruk_Click()
Dim ctrls As Integer
DEInit
With DRTransaksi
        .Hide
        Set .DataSource = rs
        .DataMember = ""
        
        With .Sections("TXData_Detail").Controls
                For ctrls = 1 To .Count
                If TypeOf .Item(ctrls) Is RptTextBox Then
                    .Item(ctrls).DataMember = ""
                    .Item(ctrls).DataField = rs.Fields(ctrls).Name
                End If
                Next
        End With
End With
DRTransaksi.Visible = True
End Sub

Private Sub menuDompet_Click()
IsiBalance.Show
End Sub

Private Sub menuPktHosting_Click()
PaketSSH.Show
End Sub

Private Sub menuStatHosting_Click()
StatusSSH.Show
End Sub

Private Sub mnLogout_Click()
Unload Me
Login.Show
End Sub

Public Sub DEInit()
getNSSH
getQueryDT
End Sub

Public Sub getQueryDT()
sql = "select * from trx where uid = '" & gSID & "' "
Set rs = bridge.Execute(sql)
End Sub
