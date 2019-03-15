Attribute VB_Name = "IntelliSense"
Option Explicit

Dim bExit As Boolean

Public Sub CariIntel(sTextBox As TextBox, sTable As String, sField As String, Optional sDBPass As String)
    On Error Resume Next
    Dim lLen As Long
    If bExit = True Or sTextBox = "" Then Exit Sub
    lLen = Len(sTextBox)
    Sql = "select * from " & sTable & " where " & sField & " like '" & sTextBox & "%'"
    Set RS = bridge.Execute(Sql)
    If RS.EOF And RS.BOF Then Exit Sub
    sTextBox.Text = RS(sField)
    If sTextBox.SelText = "" Then
        sTextBox.SelStart = lLen
    Else
        sTextBox.SelStart = InStr(sTextBox.Text, sTextBox.SelText)
    End If
    sTextBox.SelLength = Len(sTextBox.Text)
End Sub

Public Sub CheKey(lChar As Integer, sTextBox As TextBox)
    If lChar = 8 Or lChar = 46 Then 'Backspace or Delete
        bExit = True
    ElseIf lChar = 9 Or lChar = 13 Then 'Tab or Enter
        sTextBox.SelStart = Len(sTextBox)
        sTextBox.SelLength = 0
        bExit = True
    ElseIf lChar = 32 Then
        If Len(sTextBox.SelText) <> 0 Then sTextBox = sTextBox '& " "
        sTextBox.SelStart = Len(sTextBox)
        sTextBox.SelLength = 0
        bExit = True
    Else
        bExit = False
    End If
End Sub

Public Sub TxtAutoComplete(Txt As TextBox, KeyAscii As Integer)
 Dim SearchText As String
 Dim EnteredText As String
 Dim SearchLen As Long
 
 On Error GoTo ErrorHandler
 With Txt
  If .SelStart > 0 Then
   EnteredText = Left$(.Text, .SelStart)
  End If
  Select Case KeyAscii
   Case vbKeyEscape, vbKeyDelete
    .Text = vbNullString
    KeyAscii = 0
    Exit Sub
   Case vbKeyBack
    If Len(EnteredText) > 1 Then
     SearchText = LCase$(Left$(EnteredText, Len(EnteredText) - 1))
    Else
     EnteredText = vbNullString
     KeyAscii = 0
     .Text = vbNullString
     Exit Sub
    End If
   Case Else
    SearchText = LCase$(EnteredText & Chr$(KeyAscii))
  End Select
  SearchLen = Len(SearchText)
  Sql = "Select * From ssh_item_active " & " Where Left(server_cuser, " & SearchLen & ") Like " & """" & "%" & SearchText & "%" & """"
  Set RS = bridge.Execute(Sql)
  If Not RS.EOF Then
   .Text = RS.Fields(0).Value
   .SelStart = Len(SearchText)
   .SelLength = Len(.Text) - Len(SearchText)
    KeyAscii = 0
  End If
 End With
 Exit Sub
ErrorHandler:
 KeyAscii = 0
 MsgBox Err.Description
End Sub
