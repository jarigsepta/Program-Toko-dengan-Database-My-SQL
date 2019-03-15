Attribute VB_Name = "DBConnect"
Option Explicit

Public bridge As New ADODB.Connection
Public rs As New ADODB.Recordset

Public constr As String
Public sql As String

Public Sub CreateConn()
On Error GoTo Err
constr = "DRIVER={MySQL ODBC 5.3 Unicode Driver};SERVER=localhost;PORT=3306;DATABASE=OhSSH;UID=root;"
Set bridge = New ADODB.Connection
bridge.CursorLocation = adUseClient
bridge.Open (constr)
Exit Sub

Err:
   MsgBox "MySQL Koneksi Database Error..! Check Koneksi Jaringan: " & Err.Description, vbCritical, "Pesan Error"
End
End Sub
