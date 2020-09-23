VERSION 5.00
Begin VB.Form frmServer 
   Caption         =   "DDE Server"
   ClientHeight    =   1710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5160
   Icon            =   "frmServer.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "frmServer"
   ScaleHeight     =   1710
   ScaleWidth      =   5160
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDDE 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Type text here"
      Top             =   120
      Width           =   4815
   End
   Begin VB.Label lblInstructions 
      Caption         =   $"frmServer.frx":0442
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4815
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==========================================
'DDE Server
'Copyright © 2001 Steven Gilman
'------------------------------------------
'This project demonstrates use of Dynamic
'Data Exchange between two applications.
'It is for educational purposes only.
'==========================================
