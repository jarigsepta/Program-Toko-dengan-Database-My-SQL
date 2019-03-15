VERSION 5.00
Object = "{EAFDAFBF-1D88-41DD-B117-60ECBC4B8441}#1.0#0"; "vkUserControlsXP.ocx"
Begin VB.Form Agreement 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "License Agreement"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   8925
   StartUpPosition =   1  'CenterOwner
   Begin vkUserContolsXP.vkTextBox vkTextBox1 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   8281
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
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2
      LegendForeColor =   16750899
   End
End
Attribute VB_Name = "Agreement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim FSO As Object
    Dim FileName As File
    Dim TextStream As TextStream
    Dim strText As String
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If FSO.FileExists(App.Path & "\License Agreement.txt") Then
        Set FileName = FSO.GetFile(App.Path & "\License Agreement.txt")
    Else
        'Do something else
    End If
    
    Set TextStream = FileName.OpenAsTextStream(ForReading, TristateUseDefault)
        
    strText = TextStream.ReadAll
        
    Do Until TextStream.AtEndOfStream
        strText = TextStream.ReadAll
    Loop
    TextStream.Close
    vkTextBox1.Text = strText
    Set TextStream = Nothing
    Set FSO = Nothing
End Sub

