Attribute VB_Name = "getID"
Option Explicit

Public gSID As String
Public gUID As String
Public gTXID As String
Public gActiveID As String

Public Sub getNSSH()
sql = "select uid from customer where email = '" & Login.vkTextBox1.Text & "' "
Set rs = bridge.Execute(sql)
gSID = rs.Fields("uid").Value
End Sub

Public Sub getActiveIDSSH()
sql = "select item_id from ssh_item_menu where ssh_server = '" & PaketSSH.Combo1.Text & "' "
Set rs = bridge.Execute(sql)
gActiveID = rs.Fields("item_id").Value
End Sub

Public Sub getUID()
sql = "select MAX(uid) from customer"
Set rs = bridge.Execute(sql)
End Sub

Public Sub getTXID()
sql = "select tx_id from trx order by tx_id desc limit 1"
'sql = "select MAX(tx_id) from trx"
Set rs = bridge.Execute(sql)
End Sub
