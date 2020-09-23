VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvbutton.ocx"
Begin VB.Form frm_supView 
   BackColor       =   &H00CDDAB1&
   Caption         =   "View Suppliers"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5295
   ScaleWidth      =   6435
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView lstv_comV 
      Height          =   4620
      Left            =   9045
      TabIndex        =   7
      Top             =   2205
      Width           =   2730
      _ExtentX        =   4815
      _ExtentY        =   8149
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Company Name"
         Object.Width           =   4410
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      Height          =   255
      Left            =   270
      ScaleHeight     =   195
      ScaleWidth      =   270
      TabIndex        =   2
      Top             =   2535
      Width           =   330
   End
   Begin VB.TextBox txtfind 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2685
      TabIndex        =   1
      Top             =   1485
      Width           =   4560
   End
   Begin VB.TextBox txtid 
      Height          =   465
      Left            =   1995
      TabIndex        =   0
      Top             =   8460
      Width           =   1215
   End
   Begin lvButton.lvButtons_H cmdfind 
      Height          =   405
      Left            =   1710
      TabIndex        =   4
      Top             =   1485
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   714
      Caption         =   "Find"
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
   Begin MSComctlLib.ListView lstv_sup 
      Height          =   4665
      Left            =   195
      TabIndex        =   5
      Top             =   2175
      Width           =   8730
      _ExtentX        =   15399
      _ExtentY        =   8229
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483642
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Address"
         Object.Width           =   5997
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Ph no"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Height          =   780
      Left            =   30
      TabIndex        =   6
      Top             =   15
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   1376
      ButtonWidth     =   1402
      ButtonHeight    =   1376
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "itb32x32"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Shortcuts"
            Key             =   "Shortcuts"
            Object.ToolTipText     =   "Ctrl+F1"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            Key             =   "New"
            Object.ToolTipText     =   "Ctrl+F2"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit"
            Key             =   "Edit"
            Object.ToolTipText     =   "Ctrl+F3"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search"
            Key             =   "Search"
            Object.ToolTipText     =   "Ctrl+F4"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            Key             =   "Delete"
            Object.ToolTipText     =   "Ctrl+F5"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh"
            Key             =   "Refresh"
            Object.ToolTipText     =   "Ctrl+F6"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print"
            Key             =   "Print"
            Object.ToolTipText     =   "Ctrl+F7"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
            Key             =   "Close"
            Object.ToolTipText     =   "Ctrl+F8"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList itb32x32 
      Left            =   420
      Top             =   1380
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSupp_v.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSupp_v.frx":1992
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSupp_v.frx":3324
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSupp_v.frx":4CB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSupp_v.frx":6648
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSupp_v.frx":7FDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSupp_v.frx":996C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSupp_v.frx":B2FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSupp_v.frx":CC90
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSupp_v.frx":E624
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSupp_v.frx":F300
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSupp_v.frx":FBE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSupp_v.frx":108BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSupp_v.frx":11598
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSupp_v.frx":12274
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSupp_v.frx":12F50
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSupp_v.frx":13C2C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "View Total Suppliers"
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
      Left            =   165
      TabIndex        =   3
      Top             =   1035
      Width           =   1995
   End
   Begin VB.Shape shpBar 
      BackColor       =   &H009AB564&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   45
      Top             =   1020
      Width           =   11745
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00BBCD96&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00BBCD96&
      Height          =   420
      Left            =   15
      Top             =   960
      Width           =   11820
   End
End
Attribute VB_Name = "frm_supView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs_lst As New ADODB.Recordset
Dim rs_Delete As New ADODB.Recordset
Dim Rs As ADODB.Recordset
Public Sub List_sup()
If rs_lst.State = 1 Then rs_lst.Close
Dim li As ListItem
rs_lst.Open "select * from supplier", cn, adOpenDynamic, adLockReadOnly
lstv_sup.ListItems.Clear
While Not rs_lst.EOF
Set li = lstv_sup.ListItems.Add
li.Text = rs_lst!sup_id
li.SubItems(1) = rs_lst!sup_name
li.SubItems(2) = rs_lst!sup_address
'li.SubItems(3) = rs_lst!date_s
li.SubItems(3) = rs_lst!sup_ph
'li.SubItems(5) = rs_lst!sup_sex
rs_lst.MoveNext
Wend
lstv_sup.Refresh
End Sub

Private Sub Delete()
Dim a As Integer
Dim x As String
a = MsgBox("Do you really want to delete the record of " & lstv_sup.SelectedItem.SubItems(1), vbCritical + vbYesNo, "Attention")
If a = vbYes Then
If rs_Delete.State = 1 Then rs_Delete.Close
rs_Delete.ActiveConnection = cn
rs_Delete.CursorType = adOpenDynamic
rs_Delete.CursorLocation = adUseClient
rs_Delete.LockType = adLockOptimistic
x = Mid(txtid.Text, 5, Len(txtid.Text))
rs_Delete.Open "Delete  from supplier where sup_id=" & lstv_sup.SelectedItem.Text
End If
Call List_sup
End Sub

Private Sub cmdfind_Click()
txtfind.BackColor = vbWhite
txtfind.Enabled = True
txtfind.SetFocus
End Sub

Private Sub Form_Load()
Call MdItoolBar
txtfind.Enabled = False
Call SetListViewColor(lstv_sup, Picture1, vbWhite, vbGray)
Call List_sup
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call MditoolbarV
Set frm_supView = Nothing
End Sub
Private Sub lstv_sup_DblClick()
Dim li As ListItem
Dim i As Integer
With frm_suppliers
.txtid.Text = lstv_sup.SelectedItem.Text
.txtname.Text = lstv_sup.SelectedItem.SubItems(1)
.txtadd.Text = lstv_sup.SelectedItem.SubItems(2)
.txtph.Text = lstv_sup.SelectedItem.SubItems(3)
For i = 1 To .lstv_com.ListItems.Count
Set li = .lstv_com.ListItems.item(i)
Set Rs = New ADODB.Recordset '(1)
If Rs.State = 1 Then Rs.Close ' (2) these both must be in loop because on every loop it close and open recordset
Rs.Open "select com_id,company_name from qrysupcompany where sup_id=" & lstv_sup.SelectedItem.Text, cn, adOpenDynamic, adLockOptimistic
Do While Not Rs.EOF
If Rs!com_id = li.Text Then li.Checked = True ' this statement check weather there record exits in company list if it's true then show checkboxes in front of company id
Rs.MoveNext
Loop
Next
.Show
.cmdsave.Visible = False
.cmdEsave.Visible = True

End With
End Sub
Private Sub lstv_sup_ItemClick(ByVal item As MSComctlLib.ListItem)
lstv_comV.Refresh
Dim li As ListItem
Set Rs = New ADODB.Recordset
If Rs.State = 1 Then Rs.Close
Rs.Open "select company_name from qrysupcompany where sup_id =" & lstv_sup.SelectedItem.Text, cn, adOpenDynamic, adLockOptimistic
lstv_comV.ListItems.Clear
Do While Not Rs.EOF
Set li = lstv_comV.ListItems.Add
li.Text = Rs!company_name
Rs.MoveNext
Loop
lstv_comV.Refresh
End Sub

Private Sub lstv_sup_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
Call Delete
End If
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "delete"
Call Delete
Case "exit"
Unload Me
End Select
End Sub
Private Sub FindList()
Dim li As ListItem
Set Rs = New ADODB.Recordset
If Rs.State = 1 Then Set Rs = Nothing
Rs.Open "select * from supplier where sup_name like '" & Trim(txtfind) & "%'", cn, adOpenKeyset, adLockPessimistic
lstv_sup.ListItems.Clear
Do While Not Rs.EOF
Set li = lstv_sup.ListItems.Add(, , Rs!sup_id)

li.SubItems(1) = Rs!sup_name
li.SubItems(2) = Rs!sup_address
li.SubItems(3) = Rs!date_s
li.SubItems(4) = Rs!sup_ph
li.SubItems(5) = Rs!sup_sex
Rs.MoveNext
Loop
End Sub

Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "New"
frm_suppliers.Show
Case "Close"
Unload Me
End Select
End Sub

Private Sub txtfind_Change()
Call FindList
End Sub
