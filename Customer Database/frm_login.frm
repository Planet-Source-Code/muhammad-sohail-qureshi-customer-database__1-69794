VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvbutton.ocx"
Begin VB.Form frm_login 
   BackColor       =   &H00CF6154&
   BorderStyle     =   0  'None
   ClientHeight    =   5655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3780
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   3780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin lvButton.lvButtons_H cmdlogin 
      Height          =   390
      Left            =   1155
      TabIndex        =   4
      Top             =   5160
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   688
      Caption         =   "Login"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtpassword 
      Appearance      =   0  'Flat
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1020
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   4710
      Width           =   2655
   End
   Begin VB.ComboBox comuser 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1005
      TabIndex        =   2
      Top             =   4230
      Width           =   2715
   End
   Begin lvButton.lvButtons_H cmdcancel 
      Height          =   390
      Left            =   2430
      TabIndex        =   5
      Top             =   5160
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   688
      Caption         =   "Cancel"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   90
      TabIndex        =   3
      Top             =   4770
      Width           =   990
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   60
      TabIndex        =   1
      Top             =   4290
      Width           =   990
   End
   Begin VB.Image Image1 
      Height          =   4215
      Left            =   -30
      Picture         =   "frm_login.frx":0000
      Top             =   -15
      Width           =   3720
   End
End
Attribute VB_Name = "frm_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Call FrmCnt(frm_login)
End Sub

Private Sub txtpassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
MDIForm1.Show
End If
Unload Me
End Sub
