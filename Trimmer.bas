Attribute VB_Name = "Trimmer"
Option Explicit

Public Function RemoveCommaToDB(sParam As String) As String
    Dim Result As String
    Result = sParam
    Do
        Result = Replace(Result, ",", "")
    Loop Until InStr(Result, ",") = 0
    RemoveCommaToDB = Result
End Function
