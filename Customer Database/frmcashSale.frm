VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmcashSaleV 
   BackColor       =   &H00CDDAB1&
   Caption         =   "Cash Sales"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   630
      TabIndex        =   8
      Top             =   255
      Width           =   1620
   End
   Begin VB.PictureBox Picture1 
      Height          =   285
      Left            =   345
      ScaleHeight     =   225
      ScaleWidth      =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   300
   End
   Begin MSComctlLib.ListView lstv_inv 
      Height          =   6120
      Left            =   120
      TabIndex        =   1
      Top             =   1110
      Width           =   5220
      _ExtentX        =   9208
      _ExtentY        =   10795
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
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
         Text            =   "Customer Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Date"
         Object.Width           =   2117
      EndProperty
   End
   Begin MSComctlLib.ListView lst_invdet 
      Height          =   3315
      Left            =   5520
      TabIndex        =   3
      Top             =   1080
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   5847
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Item Sales"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Quantity"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Amount"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Discount Rs."
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   11535
      TabIndex        =   7
      Top             =   825
      Width           =   210
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   11445
      Top             =   810
      Width           =   300
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Detail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   5580
      TabIndex        =   6
      Top             =   825
      Width           =   3075
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E77A4B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   5520
      Top             =   810
      Width           =   6225
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Slect Invoice For Detail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   180
      TabIndex        =   5
      Top             =   885
      Width           =   3075
   End
   Begin VB.Shape shpBar 
      BackColor       =   &H00E77A4B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   105
      Top             =   870
      Width           =   5220
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "View Cash Sales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   660
      Left            =   3330
      TabIndex        =   2
      Top             =   135
      Width           =   4095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "View Cash Sales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000003&
      Height          =   660
      Left            =   3285
      TabIndex        =   0
      Top             =   210
      Width           =   4095
   End
End
Attribute VB_Name = "frmcashSaleV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs_lst As ADODB.Recordset
Private Sub Form_Load()
'Putting list shade function
Call SetListViewColor(lstv_inv, Picture1, vbWhite, vbGray)
'----------------------------------
Call List
End Sub
Private Sub lstv_inv_ItemClick(ByVal item As MSComctlLib.ListItem)
Set rs_lst = New ADODB.Recordset
If rs_lst.State = 1 Then rs_lst.Close
Text1.Text = lstv_inv.SelectedItem.Text
rs_lst.Open "select * from invoice_detail where inv_id = " & lstv_inv.SelectedItem.Text, cn, adOpenDynamic, adLockReadOnly
Dim li As ListItem
lst_invdet.ListItems.Clear
While Not rs_lst.EOF
Set li = lst_invdet.ListItems.Add
li.Text = rs_lst!item_name
li.SubItems(1) = rs_lst!quantity
li.SubItems(2) = rs_lst!amount
li.SubItems(3) = rs_lst!dis
rs_lst.MoveNext
Wend
End Sub

Private Sub List()
Dim li As ListItem
Set rs_lst = New ADODB.Recordset
If rs_lst.State = 1 Then rs_lst.Close
rs_lst.Open "select * from invoice", cn, adOpenDynamic, adLockReadOnly
lstv_inv.ListItems.Clear
While Not rs_lst.EOF
Set li = lstv_inv.ListItems.Add
With rs_lst
li.Text = !invoice_id
li.SubItems(1) = !cus_name
li.SubItems(2) = !date_i
.MoveNext
End With
Wend
End Sub
