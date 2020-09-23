VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   7005
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK I will do Coprate With You"
      Height          =   405
      Left            =   1620
      TabIndex        =   0
      Top             =   5415
      Width           =   3930
   End
   Begin VB.OLE OLE1 
      BorderStyle     =   0  'None
      Class           =   "Word.Document.8"
      Height          =   4905
      Index           =   1
      Left            =   210
      OleObjectBlob   =   "Form1.frx":0000
      TabIndex        =   1
      Top             =   225
      Width           =   6585
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

Unload Me
End Sub

