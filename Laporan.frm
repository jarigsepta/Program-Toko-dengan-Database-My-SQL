VERSION 5.00
Begin VB.Form Laporan 
   Caption         =   "Laporan Pembelian"
   ClientHeight    =   8145
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   10965
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "Laporan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim Appl As New CRAXDRT.Application
'Dim Report As New CRAXDRT.Report

Private Sub Form_Load()
'getNSSH
'ReportQuery

'Set Report = Appl.OpenReport(App.Path & "\laporan.rpt")
'CRViewer1.ReportSource = Report
'Report.DiscardSavedData
'Report.SQLQueryString = sql
'CRViewer1.ViewReport
End Sub

Private Sub Form_Resize()
'With CRViewer1
'.Top = 0
'.Left = 0
'.Width = Me.ScaleWidth
'.Height = Me.ScaleHeight
'End With
End Sub
