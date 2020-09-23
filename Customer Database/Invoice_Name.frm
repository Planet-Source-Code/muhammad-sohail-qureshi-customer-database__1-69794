VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Invoice_name 
   BackColor       =   &H00E4ACA0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invoice Name"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   5865
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   270
      Left            =   465
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   2
      Top             =   1440
      Width           =   270
   End
   Begin MSComctlLib.ListView invname_lst 
      Height          =   2655
      Left            =   180
      TabIndex        =   1
      Top             =   735
      Width           =   5370
      _ExtentX        =   9472
      _ExtentY        =   4683
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Invoice No"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Invoice Issue to"
         Object.Width           =   4762
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Double Click to Select Invoice"
      BeginProperty Font 
         Name            =   "Cooper BlkOul BT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   330
      TabIndex        =   0
      Top             =   105
      Width           =   5040
   End
End
Attribute VB_Name = "Invoice_name"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As New ADODB.Recordset
Private Sub Form_Load()
rs.ActiveConnection = cn
rs.CursorLocation = adUseClient
rs.CursorType = adOpenDynamic
rs.LockType = adLockReadOnly
rs.open "select  invoice_id,cus_name,date_c from invoice"
Call SetListViewColor(invname_lst, Picture1, vbWhite, vblightblack)
End Sub
