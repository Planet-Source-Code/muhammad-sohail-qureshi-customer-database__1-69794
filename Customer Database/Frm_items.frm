VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvbutton.ocx"
Begin VB.Form frm_Product 
   BackColor       =   &H00CDDAB1&
   Caption         =   "Add New Product Information"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "Frm_items.frx":0000
   LinkTopic       =   "Form6"
   MDIChild        =   -1  'True
   ScaleHeight     =   7140
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00BBCD96&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   4155
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H0095B165&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   105
         TabIndex        =   28
         Top             =   75
         Width           =   6870
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "View Products"
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
            Left            =   240
            TabIndex        =   29
            Top             =   45
            Width           =   2040
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CDDAB1&
      Height          =   3810
      Left            =   4380
      TabIndex        =   15
      Top             =   600
      Width           =   7485
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1125
         TabIndex        =   25
         Top             =   1275
         Width           =   885
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2250
         TabIndex        =   24
         Top             =   525
         Width           =   4500
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1155
         TabIndex        =   23
         Top             =   525
         Width           =   855
      End
      Begin VB.TextBox txtpname 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1545
         TabIndex        =   16
         Top             =   2640
         Width           =   2865
      End
      Begin MSDataListLib.DataCombo dc_company 
         Height          =   315
         Left            =   1530
         TabIndex        =   17
         Top             =   3240
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Text            =   ""
      End
      Begin VB.Label Label7 
         BackColor       =   &H00F8968B&
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name"
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
         Height          =   270
         Left            =   3015
         TabIndex        =   26
         Top             =   240
         Width           =   1665
      End
      Begin VB.Label Label6 
         BackColor       =   &H00F8968B&
         BackStyle       =   0  'Transparent
         Caption         =   "CO-ID*"
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
         Height          =   270
         Left            =   405
         TabIndex        =   22
         Top             =   570
         Width           =   780
      End
      Begin VB.Label Label2 
         BackColor       =   &H00F8968B&
         BackStyle       =   0  'Transparent
         Caption         =   "Product Name"
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
         Height          =   270
         Left            =   300
         TabIndex        =   19
         Top             =   2640
         Width           =   1155
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   315
         TabIndex        =   18
         Top             =   3105
         Width           =   1140
      End
   End
   Begin Customer.ctrlLiner ctrlLiner3 
      Height          =   30
      Left            =   60
      TabIndex        =   14
      Top             =   555
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   53
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00CDDAB1&
      Height          =   5670
      Left            =   105
      TabIndex        =   12
      Top             =   600
      Width           =   4050
      Begin VB.TextBox txtfind 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         CausesValidation=   0   'False
         ForeColor       =   &H00000080&
         Height          =   345
         Left            =   735
         TabIndex        =   30
         Top             =   990
         Width           =   3075
      End
      Begin MSComctlLib.ListView lstv_company 
         Height          =   3990
         Left            =   135
         TabIndex        =   13
         Top             =   1500
         Width           =   3780
         _ExtentX        =   6668
         _ExtentY        =   7038
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Product Name"
            Object.Width           =   4586
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   510
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Text            =   ""
      End
      Begin lvButton.lvButtons_H cmdfind 
         Height          =   375
         Left            =   105
         TabIndex        =   31
         Top             =   990
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   661
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
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   135
         TabIndex        =   20
         Top             =   255
         Width           =   1140
      End
   End
   Begin VB.TextBox txtcarriage 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9195
      TabIndex        =   10
      Top             =   4665
      Width           =   1050
   End
   Begin VB.Frame Frame7 
      Appearance      =   0  'Flat
      BackColor       =   &H00BBCD96&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   4305
      TabIndex        =   7
      Top             =   15
      Width           =   7530
      Begin VB.Frame Frame8 
         Appearance      =   0  'Flat
         BackColor       =   &H0095B165&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   105
         TabIndex        =   8
         Top             =   75
         Width           =   7305
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Add New Product"
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
            Left            =   240
            TabIndex        =   9
            Top             =   45
            Width           =   2040
         End
      End
   End
   Begin VB.TextBox txtRorder 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9240
      TabIndex        =   6
      Top             =   5220
      Width           =   1080
   End
   Begin VB.TextBox txtMin_stock 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9225
      TabIndex        =   5
      Top             =   5670
      Width           =   1080
   End
   Begin Customer.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   120
      TabIndex        =   3
      Top             =   6405
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   53
   End
   Begin lvButton.lvButtons_H cmdcancel 
      Height          =   570
      Left            =   5400
      TabIndex        =   4
      Top             =   6720
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   1005
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
      ImgAlign        =   1
      Image           =   "Frm_items.frx":000C
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdsave 
      Height          =   570
      Left            =   3675
      TabIndex        =   0
      Top             =   6720
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   1005
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
      ImgAlign        =   1
      Image           =   "Frm_items.frx":0786
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin VB.Label Label21 
      BackColor       =   &H00F8968B&
      BackStyle       =   0  'Transparent
      Caption         =   "Carriage Expense"
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
      Height          =   270
      Left            =   7575
      TabIndex        =   11
      Top             =   4710
      Width           =   1590
   End
   Begin VB.Label Label11 
      BackColor       =   &H00F8968B&
      BackStyle       =   0  'Transparent
      Caption         =   "Reorder Level"
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
      Height          =   240
      Left            =   8145
      TabIndex        =   2
      Top             =   5265
      Width           =   1110
   End
   Begin VB.Label Label20 
      BackColor       =   &H00F8968B&
      BackStyle       =   0  'Transparent
      Caption         =   "Minimum Stock"
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
      Height          =   270
      Left            =   8055
      TabIndex        =   1
      Top             =   5700
      Width           =   1095
   End
End
Attribute VB_Name = "frm_Product"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset
Dim SqlList As String
Private Sub Save_PrOduct()
On Error GoTo hell
Set Rs = New ADODB.Recordset
If Rs.State = 1 Then Rs.Close
Rs.Open "select * from product", cn, adOpenDynamic, adLockOptimistic
With Rs
.AddNew
If txtid.Text <> "" Then !product_id = txtid.Text
If txtpname.Text <> "" Then !product_name = txtpname.Text
If dc_company.Text <> "" Then !sup_id = dc_company.BoundText
If txtpur_up.Text <> "" Then !purup = txtpur_up.Text
If txtsal_up.Text <> "" Then !salup = txtsal_up.Text
If txtsal_pp.Text <> "" Then !salpp = txtsal_pp.Text
If txtkg.Text <> "" Then !kg = txtkg.Text
If txtltr.Text <> "" Then !ltr = txtltr.Text
If txtbox.Text <> "" Then !box = txtbox.Text
If txtcarton.Text <> "" Then !carton = txtcarton.Text
If txtstrip.Text <> "" Then !strip = txtstrip.Text
If txtRorder.Text <> "" Then !rorder = txtRorder.Text
If txtMin_stock.Text <> "" Then !mstock = txtMin_stock.Text
If txtcarriage.Text <> "" Then !carriage = txtcarriage.Text
.Update
End With
Exit Sub
hell:
If Err.Number = -2147217873 Then
MsgBox "Invalid Product Entery", vbOKOnly + vbCritical
End If
End Sub

Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub cmdsave_Click()
Call Save_PrOduct
End Sub
Private Sub Form_Load()
Call FrmCnt(frm_Product)
Call pRo_Id
Call MdItoolBar
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Call frm_ProductV.LiSt_Product
Call MditoolbarV
End Sub

Private Sub txtPName_GotFocus()
Call supPnAme
End Sub
Private Sub pRo_Id()
Set Rs = New ADODB.Recordset
If Rs.State = 1 Then Rs.Close
Rs.Open "select max(product_id)+1 as id from product", cn, adOpenDynamic, adLockOptimistic
If IsNull(Rs!Id) Then
txtid.Text = 5001
Else
txtid.Text = Rs!Id
End If
End Sub
Private Sub supPnAme()
Set Rs = New ADODB.Recordset
If Rs.State = 1 Then Rs.Close
Rs.Open "select com_id,company_name from company", cn, adOpenStatic, adLockOptimistic
Set dc_company.RowSource = Rs
dc_company.ListField = "company_name"
dc_company.BoundColumn = "com_id"
End Sub

