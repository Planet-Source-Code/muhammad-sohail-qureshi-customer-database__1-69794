VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvbutton.ocx"
Begin VB.Form frm_suppliers 
   BackColor       =   &H00CDDAB1&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Suppliers"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9150
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   9150
   Begin VB.PictureBox Picture1 
      Height          =   270
      Left            =   450
      ScaleHeight     =   210
      ScaleWidth      =   165
      TabIndex        =   16
      Top             =   1320
      Width           =   225
   End
   Begin VB.Frame Frame7 
      Appearance      =   0  'Flat
      BackColor       =   &H00BBCD96&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   11835
      Begin VB.Frame Frame8 
         Appearance      =   0  'Flat
         BackColor       =   &H0095B165&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   30
         TabIndex        =   14
         Top             =   75
         Width           =   11760
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Add New Suppliers"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   315
            Left            =   195
            TabIndex        =   15
            Top             =   45
            Width           =   3915
         End
      End
   End
   Begin MSComctlLib.ListView lstv_com 
      Height          =   3120
      Left            =   210
      TabIndex        =   12
      Top             =   630
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   5503
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   " ID"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Company Name"
         Object.Width           =   4322
      EndProperty
   End
   Begin VB.TextBox txtid 
      Appearance      =   0  'Flat
      CausesValidation=   0   'False
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   4785
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   675
      Width           =   1290
   End
   Begin VB.TextBox txtname 
      Appearance      =   0  'Flat
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   4785
      TabIndex        =   2
      Top             =   1155
      Width           =   2745
   End
   Begin VB.TextBox txtadd 
      Appearance      =   0  'Flat
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   4770
      TabIndex        =   1
      Top             =   1635
      Width           =   4095
   End
   Begin VB.TextBox txtph 
      Appearance      =   0  'Flat
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   4770
      TabIndex        =   0
      Top             =   2100
      Width           =   2130
   End
   Begin Customer.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   15
      TabIndex        =   4
      Top             =   4125
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   53
   End
   Begin lvButton.lvButtons_H cmdcancel 
      Height          =   585
      Left            =   4725
      TabIndex        =   5
      Top             =   4410
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1032
      Caption         =   "Cancel"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frm_Supp.frx":0000
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdsave 
      Height          =   585
      Left            =   1755
      TabIndex        =   6
      Top             =   4395
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1032
      Caption         =   "Save"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frm_Supp.frx":077A
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdEsave 
      Height          =   585
      Left            =   2850
      TabIndex        =   7
      Top             =   4410
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1032
      Caption         =   "Save"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   255
      cFHover         =   255
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frm_Supp.frx":0EF4
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00F8968B&
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier No"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   3660
      TabIndex        =   11
      Top             =   675
      Width           =   1050
   End
   Begin VB.Label Label3 
      BackColor       =   &H00F8968B&
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   3915
      TabIndex        =   10
      Top             =   1140
      Width           =   1050
   End
   Begin VB.Label Label7 
      BackColor       =   &H00F8968B&
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   3915
      TabIndex        =   9
      Top             =   1680
      Width           =   1050
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00F8968B&
      BackStyle       =   0  'Transparent
      Caption         =   "Contact No"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   3735
      TabIndex        =   8
      Top             =   2115
      Width           =   1050
   End
End
Attribute VB_Name = "frm_suppliers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset
Dim cmd As ADODB.Command
Dim Sql As String
Private Sub Save()
'On Error GoTo hell
Set Rs = New ADODB.Recordset
If Rs.State = 1 Then Rs.Close
Rs.Open "select * from supplier", cn, adOpenDynamic, adLockOptimistic
Rs.AddNew
 With Rs
 If txtid.Text <> "" Then !sup_id = txtid.Text
If txtname.Text <> "" Then !sup_name = txtname.Text
If txtadd.Text <> "" Then !sup_address = txtadd.Text
If txtph.Text <> "" Then !sup_ph = txtph.Text
Rs.Update
End With
Dim i As Integer
Dim li As ListItem
For i = 1 To lstv_com.ListItems.Count
Set li = lstv_com.ListItems.item(i)
Set Rs = New ADODB.Recordset
If Rs.State = 1 Then Rs.Close
Rs.Open "select * from supcompany", cn, adOpenDynamic, adLockOptimistic
Rs.AddNew
If li.Checked = True Then
Rs!com_id = li.Text
Rs!sup_id = txtid.Text
Rs.Update
End If
Next
'Exit Sub
'hell:
'If Err.Number = -2147217887 Then
' MsgBox "Do not Left any Field Empty", vbInformation + vbOKOnly
' End If
End Sub
Private Sub cmdcancel_Click()
Unload Me
frm_supView.lstv_comV.Refresh
frm_supView.lstv_sup.Refresh
End Sub

Private Sub cmdEsave_Click()
Sql = "delete from supcompany where sup_id= " & txtid.Text
cn.Execute Sql

 End Sub

Private Sub cmdsave_Click()
Call Save
End Sub
Private Sub Form_Load()
Call Sup_No
Call List_CoM
Call FrmCnt(frm_suppliers)
txtid.TabIndex = 0
txtname.TabIndex = 1
txtadd.TabIndex = 2
txtph.TabIndex = 3
Call SetListViewColor(lstv_com, Picture1, vbWhite, vbGray)
End Sub
Private Sub Sup_No()
Set Rs = New ADODB.Recordset
If Rs.State = 1 Then Rs.Close
Rs.Open "select max(sup_id)+1 as id from supplier", cn, adOpenDynamic, adLockOptimistic
If IsNull(Rs!Id) Then
txtid.Text = 1
Else
txtid.Text = Rs!Id
End If
End Sub
Private Sub Clear_Txt()
txtname.Text = ""
txtadd.Text = ""
txtph.Text = ""
End Sub
Private Sub Form_Unload(Cancel As Integer)
Call frm_supView.List_sup
 Set frm_suppliers = Nothing
End Sub
Private Sub List_CoM()
Dim li As ListItem
Set Rs = New ADODB.Recordset
If Rs.State = 1 Then Rs.Close
Rs.Open "select * from company", cn, adOpenDynamic, adLockReadOnly
lstv_com.ListItems.Clear
Do While Not Rs.EOF
Set li = lstv_com.ListItems.Add
li.Text = Rs!com_id
li.SubItems(1) = Rs!company_name
Rs.MoveNext
Loop
lstv_com.Refresh
End Sub

