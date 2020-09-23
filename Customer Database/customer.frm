VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvbutton.ocx"
Begin VB.Form frm_customer 
   BackColor       =   &H00CDDAB1&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10035
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "customer.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   10035
   Begin Customer.ctrlLiner ctrlLiner2 
      Height          =   30
      Left            =   15
      TabIndex        =   17
      Top             =   675
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   53
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00CDDAB1&
      Height          =   5055
      Left            =   135
      TabIndex        =   16
      Top             =   690
      Width           =   4110
      Begin VB.PictureBox Picture1 
         Height          =   255
         Left            =   450
         ScaleHeight     =   195
         ScaleWidth      =   120
         TabIndex        =   33
         Top             =   1335
         Width           =   180
      End
      Begin VB.TextBox txtfind 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         CausesValidation=   0   'False
         ForeColor       =   &H00000080&
         Height          =   360
         Left            =   810
         TabIndex        =   20
         Top             =   300
         Width           =   3045
      End
      Begin MSComctlLib.ListView lstv_custV 
         Height          =   4200
         Left            =   105
         TabIndex        =   18
         Top             =   750
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   7408
         View            =   3
         LabelEdit       =   1
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cus ID"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Customer Name"
            Object.Width           =   4410
         EndProperty
      End
      Begin lvButton.lvButtons_H cmdfind 
         Height          =   360
         Left            =   210
         TabIndex        =   19
         Top             =   300
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   635
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
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00BBCD96&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   60
      TabIndex        =   13
      Top             =   45
      Width           =   9930
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         BackColor       =   &H0095B165&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   0
         TabIndex        =   27
         Top             =   75
         Width           =   4080
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "View Customers"
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
            Height          =   345
            Left            =   180
            TabIndex        =   28
            Top             =   30
            Width           =   2100
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H0095B165&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   4725
         TabIndex        =   14
         Top             =   75
         Width           =   5085
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Add New Customers"
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
            Height          =   345
            Left            =   180
            TabIndex        =   15
            Top             =   30
            Width           =   2100
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDDAB1&
      Height          =   5085
      Left            =   4290
      TabIndex        =   1
      Top             =   690
      Width           =   5535
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   2100
         TabIndex        =   31
         Top             =   480
         Width           =   3285
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   1140
         TabIndex        =   29
         Top             =   465
         Width           =   915
      End
      Begin VB.TextBox txtid 
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1020
         Width           =   1620
      End
      Begin VB.TextBox txtnam 
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   1605
         TabIndex        =   4
         Top             =   1545
         Width           =   3345
      End
      Begin VB.TextBox txtadd 
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   1620
         TabIndex        =   3
         Top             =   2055
         Width           =   3675
      End
      Begin VB.TextBox txtph 
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   1620
         TabIndex        =   2
         Top             =   2535
         Width           =   2175
      End
      Begin lvButton.lvButtons_H cmdcancel 
         Height          =   585
         Left            =   2895
         TabIndex        =   10
         Top             =   3405
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
         Image           =   "customer.frx":000C
         ImgSize         =   24
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdsave 
         Height          =   585
         Left            =   1035
         TabIndex        =   11
         Top             =   3435
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
         Image           =   "customer.frx":0786
         ImgSize         =   24
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdEsave 
         Height          =   585
         Left            =   1050
         TabIndex        =   12
         Top             =   3420
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
         Image           =   "customer.frx":0F00
         ImgSize         =   24
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdrefresh 
         Height          =   540
         Left            =   1890
         TabIndex        =   24
         Top             =   4335
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   953
         Caption         =   "Refresh"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   1
         Image           =   "customer.frx":167A
         ImgSize         =   24
         cBack           =   -2147483633
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00F8968B&
         BackStyle       =   0  'Transparent
         Caption         =   "Area Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   3060
         TabIndex        =   32
         Top             =   225
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00F8968B&
         BackStyle       =   0  'Transparent
         Caption         =   "Area ID*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   75
         TabIndex        =   30
         Top             =   510
         Width           =   1140
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00F8968B&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer ID*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   225
         TabIndex        =   9
         Top             =   1065
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00F8968B&
         BackStyle       =   0  'Transparent
         Caption         =   "Name*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   300
         Left            =   510
         TabIndex        =   8
         Top             =   1560
         Width           =   660
      End
      Begin VB.Label Label7 
         BackColor       =   &H00F8968B&
         BackStyle       =   0  'Transparent
         Caption         =   "Address*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   405
         TabIndex        =   7
         Top             =   2115
         Width           =   930
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00F8968B&
         BackStyle       =   0  'Transparent
         Caption         =   "Contact No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   300
         Left            =   105
         TabIndex        =   6
         Top             =   2610
         Width           =   1500
      End
   End
   Begin Customer.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   60
      TabIndex        =   0
      Top             =   5835
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   53
   End
   Begin lvButton.lvButtons_H cmdnew 
      Height          =   570
      Left            =   330
      TabIndex        =   21
      Top             =   6210
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   1005
      Caption         =   "New"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   1
      Image           =   "customer.frx":1DF4
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdedit 
      Height          =   570
      Left            =   1965
      TabIndex        =   22
      Top             =   6210
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   1005
      Caption         =   "Edit"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   1
      Image           =   "customer.frx":256E
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmddelete 
      Height          =   570
      Left            =   3600
      TabIndex        =   23
      Top             =   6210
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   1005
      Caption         =   "Delete"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   1
      Image           =   "customer.frx":2CE8
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdexit 
      Height          =   660
      Left            =   8190
      TabIndex        =   25
      Top             =   6120
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   1164
      Caption         =   "Exit"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   1
      Image           =   "customer.frx":3462
      ImgSize         =   32
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdview 
      Height          =   570
      Left            =   5235
      TabIndex        =   26
      Top             =   6210
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   1005
      Caption         =   "View"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   1
      Image           =   "customer.frx":3BDC
      ImgSize         =   24
      cBack           =   -2147483633
   End
End
Attribute VB_Name = "frm_customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset
Private Sub Save()
On Error GoTo hell
cn.BeginTrans
Set Rs = New ADODB.Recordset
If Rs.State = 1 Then Rs.Close
Rs.Open "select * from customer", cn, adOpenDynamic, adLockOptimistic
Rs.AddNew
 With Rs
If txtid.Text <> "" Then !cus_id = txtid.Text
If txtnam.Text <> "" Then !name_c = txtnam.Text
If txtadd.Text <> "" Then !address = txtadd.Text
If txtph.Text <> "" Then !ph = txtph.Text
.Update
End With
cn.CommitTrans
MsgBox "Record Save Sucessfully", vbOKOnly, "Save"
If vbOK Then TxtClear: TeXtEnable
Exit Sub
hell:
cn.RollbackTrans
MsgBox "Error in data Saving Check then click to Save", Err.Number = -2147217887, "Soft Vision"
End Sub
Private Sub TxtClear()
Call Cus_No
txtnam.SetFocus
txtnam.Text = ""
txtadd.Text = ""
txtph.Text = ""
End Sub
Private Sub cmdcancel_Click()
'Rs.CancelUpdate

End Sub

Private Sub cmdedit_Click()
TeXtEnableT
txtnam.SetFocus
cmdsave.Visible = False
cmdEsave.Visible = True
End Sub

Private Sub cmdEsave_Click()
Set Rs = New ADODB.Recordset
If Rs.State = 1 Then Rs.Close
Rs.Open "update customer set name_c='" & txtnam.Text & "',address='" & txtadd.Text & "', ph='" & txtph.Text & "' where cus_id = " & txtid.Text, cn, adOpenDynamic, adLockOptimistic
MsgBox "Record Edit Successfully", vbInformation + vbOKOnly, "Edited"
If vbYes Then Call TeXtEnable
End Sub

Private Sub cmdexit_Click()
Unload Me

End Sub

Private Sub cmdfind_Click()
txtfind.Enabled = True
txtfind.SetFocus
txtfind.BackColor = vbWhite
End Sub

Private Sub cmdnew_Click()
Call TeXtEnableT
Cus_No
txtnam.SetFocus
cmdsave.Enabled = True
End Sub

Private Sub cmdrefresh_Click()
Call List_customerView
End Sub

Private Sub cmdsave_Click()
Call Save
List_customerView
End Sub

Private Sub cmdview_Click()
Frame4.Enabled = True
lstv_custV.ForeColor = vbBlack
cmdedit.Enabled = True
cmddelete.Enabled = True
End Sub

Private Sub Form_Load()
Cmd_enAbledF 'enabled false the buttons save,edit,delete
Call TeXtEnable
'Call Cus_No
Call FrmCnt(frm_customer)
txtid.TabIndex = 0
txtnam.TabIndex = 1
txtadd.TabIndex = 2
txtph.TabIndex = 3
cmdsave.TabIndex = 4
Call List_customerView
Call SetListViewColorShort(lstv_custV, Picture1, vbWhite, vbGray)
End Sub
Private Sub Cus_No() 'generate the custoemr id no
Set Rs = New ADODB.Recordset
If Rs.State = 1 Then Rs.Close
Rs.Open "select max(cus_id)+1 as id from customer", cn, adOpenDynamic, adLockOptimistic
If IsNull(Rs!Id) Then
txtid.Text = 1
Else
txtid.Text = Rs!Id
End If
End Sub
Private Sub Form_Unload(Cancel As Integer) 'this function make null value of form when we unload it
 Set frm_customer = Nothing
End Sub

Private Sub cmddelete_Click() 'To delete any record
Set Rs = New ADODB.Recordset
Dim a As Integer
a = MsgBox("Do you really want to delete the record of " & lstv_custV.SelectedItem.SubItems(1), vbCritical + vbYesNo, "Attention")
If a = vbYes Then
If Rs.State = 1 Then Rs.Close
Rs.ActiveConnection = cn
Rs.CursorType = adOpenDynamic
Rs.CursorLocation = adUseClient
Rs.LockType = adLockOptimistic
Rs.Open "Delete  from customer where cus_id= " & lstv_custV.SelectedItem.Text
End If
lstv_custV.Refresh
List_customerView
End Sub

Private Sub lstv_custV_ItemClick(ByVal item As MSComctlLib.ListItem) 'This function shows the text in textboxes on which we click
Set Rs = New ADODB.Recordset
If Rs.State = 1 Then Rs.Close
Rs.Open "select * from customer where cus_id=" & lstv_custV.SelectedItem.Text, cn, adOpenDynamic, adLockOptimistic
txtid.Text = Rs!cus_id
txtnam.Text = Rs!name_c
txtadd.Text = Rs!address
txtph.Text = Rs!ph
If txtnam.Enabled = True Then txtnam.SetFocus
End Sub

Private Sub txtadd_GotFocus()
Call TextGotF(txtadd)
End Sub

Private Sub txtadd_KeyPress(KeyAscii As Integer)
Call TextValid(KeyAscii, 2)
Call Pcase(KeyAscii, txtadd)
End Sub

Private Sub txtadd_KeyUp(KeyCode As Integer, Shift As Integer)
Call TextKeyD(KeyCode)
End Sub

Private Sub txtadd_LostFocus()
Call TextLostF(txtadd)
End Sub

Private Sub txtfind_Change()
FindList
End Sub

Private Sub txtnam_GotFocus()
 Call TextGotF(txtnam)
End Sub

Private Sub txtnam_KeyPress(KeyAscii As Integer)
Call TextValid(KeyAscii, 2)
Call Pcase(KeyAscii, txtnam)

End Sub

Private Sub txtnam_KeyUp(KeyCode As Integer, Shift As Integer)
Call TextKeyD(KeyCode)
End Sub

Private Sub txtnam_LostFocus()
Call TextLostF(txtnam)
End Sub

Private Sub txtph_GotFocus()
Call TextGotF(txtph)
End Sub

Private Sub txtph_KeyPress(KeyAscii As Integer)
Call TextValid(KeyAscii, 1)
If KeyAscii = 13 Then Call cmdsave_Click
End Sub

Private Sub txtph_KeyUp(KeyCode As Integer, Shift As Integer)
Call TextKeyD(KeyCode)
End Sub
Private Sub txtph_LostFocus()
Call TextLostF(txtph)
End Sub
Private Sub List_customerView() 'this function view all customer that will in database
Dim li As ListItem
Set Rs = New ADODB.Recordset
If Rs.State = 1 Then Rs.Close
Rs.Open "select cus_id,name_c from customer", cn, adOpenDynamic, adLockReadOnly
lstv_custV.ListItems.Clear
Do While Not Rs.EOF
Set li = lstv_custV.ListItems.Add
li.Text = Rs!cus_id
li.SubItems(1) = Rs!name_c
Rs.MoveNext
Loop
End Sub
Private Sub FindList() 'this function used to find any text from listview
Dim li As ListItem
Set Rs = New ADODB.Recordset
If Rs.State = 1 Then Set Rs = Nothing
Rs.Open "select * from customer where name_c like '" & Trim(txtfind) & "%'", cn, adOpenKeyset, adLockPessimistic
lstv_custV.ListItems.Clear
Do While Not Rs.EOF
Set li = lstv_custV.ListItems.Add(, , Rs!cus_id)
li.SubItems(1) = Rs!name_c
Rs.MoveNext
Loop
End Sub
Private Sub TeXtEnable() 'Enabled false of all textbox in form
lstv_custV.ForeColor = &HC0C0C0
Frame4.Enabled = False
Dim a As Variant
For Each a In frm_customer
If TypeOf a Is TextBox Then
a.Enabled = False
a.BackColor = &HC0FFFF
End If
Next a
End Sub
Private Sub TeXtEnableT() 'Enabled True of all textbox in form
'lstv_custV.Enabled = True
Dim a As Variant
For Each a In frm_customer
If TypeOf a Is TextBox Then
a.Enabled = True
a.BackColor = vbWhite
txtfind.BackColor = &HC0FFFF
End If
Next a
End Sub
Private Sub Cmd_enAbledF() 'It will enabled false ther command buttons
cmdedit.Enabled = False
cmddelete.Enabled = False
cmdsave.Enabled = False
End Sub
Private Sub TextClear() 'this function will clear all text and chage color
Dim a As Variant
For Each a In frm_customer
If TypeOf a Is TextBox Then
a.Text = ""
a.BackColor = &HC0FFFF
End If
Next a
End Sub

