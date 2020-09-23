VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_productinf 
   BackColor       =   &H00CDDAB1&
   Caption         =   "Product Information"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00BBCD96&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   -15
      ScaleHeight     =   465
      ScaleWidth      =   11880
      TabIndex        =   15
      Top             =   2970
      Width           =   11880
      Begin VB.PictureBox Picture3 
         BackColor       =   &H009AB564&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   -105
         ScaleHeight     =   285
         ScaleWidth      =   11910
         TabIndex        =   16
         Top             =   75
         Width           =   11910
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "View Product Information"
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
            TabIndex        =   17
            Top             =   15
            Width           =   4635
         End
      End
   End
   Begin VB.TextBox txtQP_Packing 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   7140
      TabIndex        =   12
      Top             =   2505
      Width           =   1080
   End
   Begin VB.ComboBox comS_packing 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3270
      TabIndex        =   11
      Top             =   2505
      Width           =   1335
   End
   Begin VB.TextBox txtMin_stock 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   5490
      TabIndex        =   9
      Top             =   2505
      Width           =   1080
   End
   Begin VB.ComboBox comP_pack 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1665
      TabIndex        =   7
      Top             =   2505
      Width           =   1335
   End
   Begin VB.TextBox txtpu_price 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   300
      TabIndex        =   5
      Top             =   2445
      Width           =   1080
   End
   Begin VB.ComboBox comproduct 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5430
      TabIndex        =   4
      Top             =   1545
      Width           =   5295
   End
   Begin VB.ComboBox comsupplier 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   135
      TabIndex        =   2
      Top             =   1515
      Width           =   4590
   End
   Begin VB.PictureBox Picture25 
      BackColor       =   &H00BBCD96&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   -150
      ScaleHeight     =   465
      ScaleWidth      =   11880
      TabIndex        =   0
      Top             =   675
      Width           =   11880
      Begin VB.PictureBox Picture2 
         BackColor       =   &H009AB564&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   -105
         ScaleHeight     =   285
         ScaleWidth      =   11910
         TabIndex        =   18
         Top             =   90
         Width           =   11910
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Define Product Information"
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
            Left            =   315
            TabIndex        =   19
            Top             =   0
            Width           =   4635
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   630
      Top             =   4260
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   23
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_produsctinf.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_produsctinf.frx":1992
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_produsctinf.frx":3324
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_produsctinf.frx":3A9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_produsctinf.frx":4218
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_produsctinf.frx":4992
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_produsctinf.frx":510C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_produsctinf.frx":5886
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_produsctinf.frx":7218
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_produsctinf.frx":8BAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_produsctinf.frx":A53C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_produsctinf.frx":BECE
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_produsctinf.frx":D860
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_produsctinf.frx":F1F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_produsctinf.frx":10B86
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_produsctinf.frx":11862
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_produsctinf.frx":12142
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_produsctinf.frx":12E1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_produsctinf.frx":13AFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_produsctinf.frx":147D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_produsctinf.frx":154B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_produsctinf.frx":1618E
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_produsctinf.frx":16A6A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Qty Per Packing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   6990
      TabIndex        =   14
      Top             =   2250
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Packing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3285
      TabIndex        =   13
      Top             =   2190
      Width           =   1275
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Minimum Stock Level"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5055
      TabIndex        =   10
      Top             =   2265
      Width           =   1845
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Packing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1635
      TabIndex        =   8
      Top             =   2175
      Width           =   1605
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Per Unit Price"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   210
      TabIndex        =   6
      Top             =   2190
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Products"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5520
      TabIndex        =   3
      Top             =   1305
      Width           =   885
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier/Agent/Company Name "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   180
      TabIndex        =   1
      Top             =   1260
      Width           =   2940
   End
End
Attribute VB_Name = "frm_productinf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset
Private Sub com_SupplierName()
Set Rs = New ADODB.Recordset
If Rs.State = 1 Then Rs.Close
Rs.Open "select sup_name from supplier", cn, adOpenDynamic, adLockOptimistic
Do While Not Rs.EOF
comsupplier.AddItem Rs!sup_name
Rs.MoveNext
Loop
End Sub

Private Sub Check1_Click()
fme_packing.Enabled = True
End Sub

Private Sub comP_pack_GotFocus()
Call com_Packing
End Sub
Private Sub comsupplier_Click()
Call com_ProductName
End Sub
Private Sub Form_Load()
Call com_SupplierName
End Sub
Private Sub com_ProductName()
Set Rs = New ADODB.Recordset
If Rs.State = 1 Then Rs.Close
Rs.Open "select product_name from product where product_supplier like '" & comsupplier.Text & "'", cn, adOpenDynamic, adLockOptimistic
comproduct.Clear
If Rs.EOF = True Then MsgBox "Supplier has no Product"
Do While Not Rs.EOF
comproduct.AddItem Rs!product_name
Rs.MoveNext
Loop
End Sub


