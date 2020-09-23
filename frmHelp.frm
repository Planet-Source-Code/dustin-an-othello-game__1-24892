VERSION 5.00
Begin VB.Form frmHelp 
   BackColor       =   &H00000000&
   Caption         =   "Help"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdok 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   6000
      Width           =   1095
   End
   Begin VB.OLE OLE2 
      AutoActivate    =   0  'Manual
      AutoVerbMenu    =   0   'False
      BackStyle       =   0  'Transparent
      Class           =   "Word.Document.8"
      Height          =   7695
      Left            =   6600
      OleObjectBlob   =   "frmHelp.frx":020A
      SourceDoc       =   "C:\Dustin\rules.doc"
      TabIndex        =   1
      Top             =   120
      Width           =   5175
   End
   Begin VB.OLE OLE1 
      AutoActivate    =   0  'Manual
      AutoVerbMenu    =   0   'False
      BackStyle       =   0  'Transparent
      Class           =   "Word.Document.8"
      Height          =   7695
      Left            =   120
      OleObjectBlob   =   "frmHelp.frx":2E022
      SourceDoc       =   "C:\Dustin\rules.doc"
      TabIndex        =   0
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
frmHelp.Hide
End Sub
