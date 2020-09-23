VERSION 5.00
Begin VB.Form frmClient 
   Caption         =   "DDE Client"
   ClientHeight    =   2145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmClient.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2145
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   $"frmClient.frx":0442
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4455
   End
   Begin VB.Label lblDDE 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   120
      LinkItem        =   "txtDDE"
      LinkTopic       =   "DDESERV|frmServer"
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==========================================
'DDE Client
'Copyright © 2001 Steven Gilman
'------------------------------------------
'This project demonstrates use of Dynamic
'Data Exchange between two applications.
'It is for educational purposes only.
'==========================================

Private Sub Form_Load()
    '
    'Enable DDE
    lblDDE.LinkMode = 1
    '
End Sub

