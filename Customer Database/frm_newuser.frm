VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvbutton.ocx"
Begin VB.Form frm_newuser 
   BackColor       =   &H0092DAE0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New User"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4380
   FillColor       =   &H0080C0FF&
   Icon            =   "frm_newuser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   4380
   StartUpPosition =   3  'Windows Default
   Begin lvButton.lvButtons_H cmdok 
      Height          =   360
      Left            =   1110
      TabIndex        =   6
      Top             =   1995
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   635
      Caption         =   "OK"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
   Begin VB.TextBox txtcpass 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1620
      PasswordChar    =   "l"
      TabIndex        =   5
      Top             =   1395
      Width           =   1965
   End
   Begin VB.TextBox txtpass 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1635
      PasswordChar    =   "l"
      TabIndex        =   3
      Top             =   825
      Width           =   1965
   End
   Begin VB.TextBox txtuser 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   1635
      TabIndex        =   1
      Top             =   255
      Width           =   1965
   End
   Begin lvButton.lvButtons_H cmdcancel 
      Height          =   360
      Left            =   2175
      TabIndex        =   7
      Top             =   1995
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   635
      Caption         =   "Cancel"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   30
      TabIndex        =   4
      Top             =   1470
      Width           =   1515
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   405
      TabIndex        =   2
      Top             =   930
      Width           =   960
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   405
      TabIndex        =   0
      Top             =   345
      Width           =   960
   End
End
Attribute VB_Name = "frm_newuser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Rs As New ADODB.Recordset

Private Sub cmdok_Click()
If txtpass.Text = txtcpass.Text Then
Rs.AddNew
Rs!user_id = txtuser.Text
Rs!pass = txtpass.Text
Rs.Update
MsgBox "New user added successfully", vbOKOnly + vbInformation, "New User"
txtpass.Text = ""
txtcpass.Text = ""
txtuser.Text = ""
Unload Me
Else
MsgBox "Password Field not matches,Try Again", vbOKOnly + vbCritical, "Allert"
txtpass.Text = ""
txtcpass.Text = ""
txtpass.SetFocus
Exit Sub
End If
End Sub

Private Sub Form_Load()
Call FrmCnt(frm_newuser)
Rs.Open "select * from pass", cn, adOpenDynamic, adLockOptimistic
If Rs.BOF And Rs.EOF = False Then
txtuser.Text = Rs!user_id
txtpass.Text = Rs!pass
End If
End Sub

