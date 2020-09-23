VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvbutton.ocx"
Begin VB.Form frm_purchase 
   BackColor       =   &H00CDDAB1&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   11715
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNamount 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9915
      Locked          =   -1  'True
      TabIndex        =   56
      Top             =   6345
      Width           =   1545
   End
   Begin VB.TextBox txtdiscR 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9930
      TabIndex        =   54
      Top             =   5970
      Width           =   1545
   End
   Begin VB.TextBox txtGtotal 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9915
      Locked          =   -1  'True
      TabIndex        =   52
      Top             =   5625
      Width           =   1545
   End
   Begin Customer.ctrlLiner ctrlLiner3 
      Height          =   30
      Left            =   30
      TabIndex        =   47
      Top             =   6840
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   53
   End
   Begin MSDataListLib.DataCombo dc_pname 
      Height          =   315
      Left            =   135
      TabIndex        =   46
      Top             =   2370
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
   End
   Begin VB.Frame frme_unitc 
      Appearance      =   0  'Flat
      BackColor       =   &H0095B165&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2955
      Left            =   120
      TabIndex        =   39
      Top             =   2745
      Visible         =   0   'False
      Width           =   6075
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1725
         TabIndex        =   40
         Top             =   2565
         Width           =   3615
      End
      Begin lvButton.lvButtons_H cmdfhide 
         Height          =   285
         Left            =   5595
         TabIndex        =   41
         Top             =   60
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   503
         Caption         =   "X"
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
         cFore           =   -2147483628
         cFHover         =   16777215
         cBhover         =   255
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   8421631
      End
      Begin MSComctlLib.ListView lstv_ucost 
         Height          =   2130
         Left            =   105
         TabIndex        =   42
         Top             =   375
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   3757
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
         Appearance      =   0
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Product Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Unit cost"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Actual Unit Cost"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Unit Cost(Sale)"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Find Product Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   285
         TabIndex        =   44
         Top             =   2580
         Width           =   1410
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Double Click For Select any Product"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   435
         TabIndex        =   43
         Top             =   90
         Width           =   4740
      End
   End
   Begin MSComctlLib.ListView lstv_Purchase 
      Height          =   2160
      Left            =   105
      TabIndex        =   38
      Top             =   3255
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   3810
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ColHdrIcons     =   "ImageList1"
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
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Product Name"
         Object.Width           =   4939
         ImageIndex      =   1
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Unit Cost"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Kg"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Box"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Carton"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Strip"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Pieces"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Total Qty"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "Total Amount"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "Purchase No"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Text            =   "Product ID"
         Object.Width           =   2293
      EndProperty
   End
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   315
      Left            =   2505
      TabIndex        =   37
      Top             =   2370
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
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
      ImgAlign        =   4
      Image           =   "frm_purchase.frx":0000
      cBack           =   -2147483633
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H0095B165&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   30
      TabIndex        =   35
      Top             =   2835
      Width           =   11760
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase Products"
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
         Left            =   165
         TabIndex        =   36
         Top             =   45
         Width           =   2040
      End
   End
   Begin lvButton.lvButtons_H cmdpurchase 
      Height          =   330
      Left            =   10725
      TabIndex        =   33
      Top             =   2325
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   582
      Caption         =   "Purchase"
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
   Begin VB.TextBox txtTamount 
      Appearance      =   0  'Flat
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
      Height          =   285
      Left            =   9570
      TabIndex        =   32
      Top             =   2355
      Width           =   1065
   End
   Begin VB.TextBox txtTqty 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8505
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   2370
      Width           =   960
   End
   Begin VB.TextBox txtkg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3810
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   2370
      Width           =   660
   End
   Begin VB.TextBox txtbox 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4515
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   2370
      Width           =   660
   End
   Begin VB.TextBox txtcarton 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5250
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   2370
      Width           =   660
   End
   Begin VB.TextBox txtstrip 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5985
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   2370
      Width           =   660
   End
   Begin VB.TextBox txtpieces 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6690
      TabIndex        =   20
      Top             =   2370
      Width           =   660
   End
   Begin VB.TextBox txtpup 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2955
      TabIndex        =   18
      Top             =   2370
      Width           =   765
   End
   Begin Customer.ctrlLiner ctrlLiner2 
      Height          =   30
      Left            =   -75
      TabIndex        =   15
      Top             =   2055
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   53
   End
   Begin VB.ComboBox compType 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9375
      Style           =   1  'Simple Combo
      TabIndex        =   14
      Top             =   1560
      Width           =   2280
   End
   Begin VB.TextBox txtsupph 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0F4F3&
      Height          =   285
      Left            =   9885
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   975
      Width           =   1680
   End
   Begin VB.TextBox txtsup_address 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0F4F3&
      Height          =   285
      Left            =   8610
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   660
      Width           =   2925
   End
   Begin Customer.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   15
      TabIndex        =   8
      Top             =   1395
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   53
   End
   Begin VB.TextBox txtid 
      Appearance      =   0  'Flat
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
      Height          =   285
      Left            =   1110
      TabIndex        =   3
      Top             =   660
      Width           =   1995
   End
   Begin VB.Frame Frame7 
      Appearance      =   0  'Flat
      BackColor       =   &H00BBCD96&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   -45
      TabIndex        =   0
      Top             =   30
      Width           =   11835
      Begin VB.Frame Frame8 
         Appearance      =   0  'Flat
         BackColor       =   &H0095B165&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   15
         TabIndex        =   1
         Top             =   45
         Width           =   11760
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Purchases"
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
            Left            =   165
            TabIndex        =   2
            Top             =   45
            Width           =   2040
         End
      End
   End
   Begin MSDataListLib.DataCombo dc_supplier 
      Height          =   315
      Left            =   1200
      TabIndex        =   4
      Top             =   1560
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker dp_purchase 
      Height          =   285
      Left            =   1095
      TabIndex        =   45
      Top             =   975
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   503
      _Version        =   393216
      CalendarBackColor=   16570065
      CalendarForeColor=   -2147483647
      CalendarTitleBackColor=   -2147483639
      CalendarTitleForeColor=   -2147483641
      CalendarTrailingForeColor=   -2147483635
      Format          =   20643841
      CurrentDate     =   39231
   End
   Begin lvButton.lvButtons_H cmdcancel 
      Height          =   555
      Left            =   7185
      TabIndex        =   48
      Top             =   7035
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   979
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
      Image           =   "frm_purchase.frx":0419
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdsave 
      Height          =   555
      Left            =   4485
      TabIndex        =   49
      Top             =   7035
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   979
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
      Image           =   "frm_purchase.frx":0B93
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5325
      Top             =   495
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchase.frx":130D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin lvButton.lvButtons_H cmdrefresh 
      Height          =   555
      Left            =   3135
      TabIndex        =   58
      Top             =   7035
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   979
      Caption         =   "New"
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
      Image           =   "frm_purchase.frx":2C9F
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmddelete 
      Height          =   555
      Left            =   5835
      TabIndex        =   59
      Top             =   7035
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   979
      Caption         =   "Delete"
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
      Image           =   "frm_purchase.frx":3419
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin VB.Label lablactTotal 
      BackColor       =   &H00F8968B&
      BackStyle       =   0  'Transparent
      Caption         =   "Net Amount"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Left            =   8820
      TabIndex        =   57
      Top             =   6345
      Width           =   1050
   End
   Begin VB.Label Label21 
      BackColor       =   &H00F8968B&
      BackStyle       =   0  'Transparent
      Caption         =   "Discount Rs."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Left            =   8820
      TabIndex        =   55
      Top             =   6030
      Width           =   1050
   End
   Begin VB.Label Label20 
      BackColor       =   &H00F8968B&
      BackStyle       =   0  'Transparent
      Caption         =   "Gross Total"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Left            =   8820
      TabIndex        =   53
      Top             =   5655
      Width           =   1050
   End
   Begin VB.Label Label19 
      BackColor       =   &H00F8968B&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Qty"
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
      Left            =   90
      TabIndex        =   51
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label14 
      BackColor       =   &H00F8968B&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Qty"
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
      Left            =   90
      TabIndex        =   50
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label16 
      BackColor       =   &H0095B165&
      Height          =   285
      Left            =   7410
      TabIndex        =   34
      Top             =   2370
      Width           =   1050
   End
   Begin VB.Label Label15 
      BackColor       =   &H00F8968B&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Left            =   9750
      TabIndex        =   31
      Top             =   2145
      Width           =   735
   End
   Begin VB.Label Label13 
      BackColor       =   &H00F8968B&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Qty"
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
      Left            =   8595
      TabIndex        =   29
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label11 
      BackColor       =   &H00F8968B&
      BackStyle       =   0  'Transparent
      Caption         =   "Kg"
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
      Height          =   285
      Left            =   3960
      TabIndex        =   27
      Top             =   2160
      Width           =   660
   End
   Begin VB.Label Label9 
      BackColor       =   &H00F8968B&
      BackStyle       =   0  'Transparent
      Caption         =   "Box"
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
      Height          =   285
      Left            =   4710
      TabIndex        =   25
      Top             =   2160
      Width           =   660
   End
   Begin VB.Label Label8 
      BackColor       =   &H00F8968B&
      BackStyle       =   0  'Transparent
      Caption         =   "Carton"
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
      Height          =   285
      Left            =   5295
      TabIndex        =   23
      Top             =   2160
      Width           =   660
   End
   Begin VB.Label Label7 
      BackColor       =   &H00F8968B&
      BackStyle       =   0  'Transparent
      Caption         =   "Strip"
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
      Height          =   285
      Left            =   6105
      TabIndex        =   21
      Top             =   2160
      Width           =   660
   End
   Begin VB.Label Label6 
      BackColor       =   &H00F8968B&
      BackStyle       =   0  'Transparent
      Caption         =   "Pieces"
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
      Height          =   285
      Left            =   6780
      TabIndex        =   19
      Top             =   2160
      Width           =   660
   End
   Begin VB.Label Label3 
      BackColor       =   &H00F8968B&
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Price"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Left            =   2940
      TabIndex        =   17
      Top             =   2160
      Width           =   810
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   180
      TabIndex        =   16
      Top             =   2130
      Width           =   1170
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Type"
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
      Left            =   8280
      TabIndex        =   13
      Top             =   1785
      Width           =   1170
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Contect No"
      ForeColor       =   &H0000011D&
      Height          =   240
      Index           =   1
      Left            =   8565
      TabIndex        =   12
      Top             =   990
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier Address"
      ForeColor       =   &H0000011D&
      Height          =   240
      Index           =   2
      Left            =   7290
      TabIndex        =   10
      Top             =   705
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   240
      Index           =   0
      Left            =   405
      TabIndex        =   7
      Top             =   1005
      Width           =   600
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase From"
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
      Left            =   105
      TabIndex        =   6
      Top             =   1620
      Width           =   1170
   End
   Begin VB.Label Label4 
      BackColor       =   &H00F8968B&
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase No"
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
      Left            =   120
      TabIndex        =   5
      Top             =   690
      Width           =   1080
   End
End
Attribute VB_Name = "frm_purchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset
Dim SQLFrame As String, SqlProBysup As String
Dim DubKg As Double, dubbox As Double, dubcarton As Double, dubstrip As Double
Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub cmdfhide_Click()
frme_unitc.Visible = False
End Sub

Private Sub cmdpurchase_Click()
'---- It will force to user to select any product
If dc_pname = "" Then
MsgBox "Please Select a Product", vbOKOnly + vbInformation, "Error "
Exit Sub
End If
'-----it will force the user to define qty of product
If txtTamount.Text = "" Or txtTamount.Text = 0 Then
MsgBox "Please Put Qty of Purchase !", vbOKOnly + vbInformation, "Error"
Exit Sub
End If
'------------

'--------------

Dim li As ListItem
Set li = lstv_Purchase.ListItems.Add

li.Text = dc_pname.Text
li.SubItems(1) = txtpup.Text
li.SubItems(2) = txtkg.Text
li.SubItems(3) = txtbox.Text
li.SubItems(4) = txtcarton.Text
li.SubItems(5) = txtstrip.Text
li.SubItems(6) = txtpieces.Text
li.SubItems(7) = txtTqty.Text
li.SubItems(8) = txtTamount.Text
li.SubItems(9) = txtid.Text
li.SubItems(10) = dc_pname.BoundText
'-------Adding Zero Value in Null Text On The Listview
If li.SubItems(2) = "" Then li.SubItems(2) = 0
If li.SubItems(3) = "" Then li.SubItems(3) = 0
If li.SubItems(4) = "" Then li.SubItems(4) = 0
If li.SubItems(5) = "" Then li.SubItems(5) = 0
If li.SubItems(6) = "" Then li.SubItems(6) = 0
Call GTotal_Amount
Call ClearAll
End Sub

Private Sub cmdsave_Click()
If txtdiscR.Text > txtGtotal.Text Then
MsgBox "Discount Rate is very High ", vbCritical + vbOKOnly, "Discount Error"
End If
End Sub

Private Sub dc_supplier_Change()
Call pRoductBysuPP
Call productCoMbo
End Sub
Private Sub Form_Load()
Call FrmCnt(frm_purchase)
dp_purchase.Value = Date
Call dc_suPr
Call Pur_iD
cmdsave.Enabled = False
cmddelete.Enabled = False
End Sub
Private Sub dc_suPr()
Set Rs = New ADODB.Recordset
If Rs.State = 1 Then Rs.Close
Rs.Open "select sup_id,sup_name from supplier", cn, adOpenStatic, adLockOptimistic
Set dc_supplier.RowSource = Rs
dc_supplier.ListField = "sup_name"
dc_supplier.BoundColumn = "sup_id"
End Sub
Private Sub lstv_Purchase_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
Dim li As ListItem
Set li = lstv_Purchase.SelectedItem
'On Error GoTo hell
Dim y As Integer
y = MsgBox("Are You Sure to Delete This Product", vbYesNo + vbInformation + vbDefaultButton2, "Information")
If y = vbYes Then
txtGtotal.Text = Round(CSng(txtGtotal.Text) - CSng(li.SubItems(8)), 2)
lstv_Purchase.ListItems.Remove li.Index
End If
End If
Exit Sub
End Sub

Private Sub lstv_ucost_DblClick()
Dim li As ListItem
Set Rs = New ADODB.Recordset
If Rs.State = 1 Then Rs.Close
Rs.Open "select * from qrypurchaseframe where product_name='" & lstv_ucost.SelectedItem.Text & "'", cn, adOpenDynamic, adLockOptimistic
txtpup.Text = ""
dc_pname.Text = lstv_ucost.SelectedItem.Text
txtpup.Text = lstv_ucost.SelectedItem.SubItems(2)
'-------------------------------
If Rs!kg <= 0 Then
txtkg.Locked = True
Call ClsTextBox
txtkg.Text = ""
txtkg.BackColor = &HD0F4F3
End If
If Rs!kg > 0 Then
Call ClsTextBox
DubKg = Val(Rs!kg)
txtkg.Locked = False
txtkg.BackColor = vbWhite
End If
'-------

If Rs!box <= 0 Then
Call ClsTextBox
txtbox.Text = ""
txtbox.Locked = True
txtbox.BackColor = &HD0F4F3
End If
If Rs!box > 0 Then
Call ClsTextBox
dubbox = Val(Rs!box)
txtbox.Locked = False
txtbox.BackColor = vbWhite
End If
'---------------------
If Rs!carton <= 0 Then
Call ClsTextBox
txtcarton.Text = ""
txtcarton.Locked = True
txtcarton.BackColor = &HD0F4F3
End If
If Rs!carton > 0 Then
Call ClsTextBox
dubcarton = Val(Rs!carton)
txtcarton.Locked = False
txtcarton.BackColor = vbWhite
End If
'----------------
If Rs!strip <= 0 Then
Call ClsTextBox
txtstrip.Text = ""
txtstrip.Locked = True
txtstrip.BackColor = &HD0F4F3
End If
If Rs!strip > 0 Then
Call ClsTextBox
 dubstrip = Val(Rs!strip)
txtstrip.Locked = False
txtstrip.BackColor = vbWhite
End If
End Sub
Private Sub lvButtons_H1_Click()
frme_unitc.Visible = True
dc_supplier.Enabled = False
End Sub
Private Sub Pur_iD()
Set Rs = New ADODB.Recordset
If Rs.State = 1 Then Rs.Close
Rs.Open "select max(purchase_id)+1 as id from purchase", cn, adOpenDynamic, adLockOptimistic
If IsNull(Rs!Id) Then
txtid.Text = 10001
Else
txtid.Text = Rs!Id
End If
End Sub
Private Sub pRoductBysuPP()
On Error GoTo hell
Dim li As ListItem
Set Rs = New ADODB.Recordset
If Rs.State = 1 Then Rs.Close
SqlProBysup = "select * from qryproductbysupplier where sup_name ='" & dc_supplier.Text & "'"
Rs.Open SqlProBysup, cn, adOpenDynamic, adLockReadOnly
txtsup_address.Text = Rs!sup_address
txtsupph.Text = Rs!sup_ph
lstv_ucost.ListItems.Clear
Do While Not Rs.EOF
Set li = lstv_ucost.ListItems.Add
li.Text = Rs!product_name
li.SubItems(1) = Rs!purup
li.SubItems(2) = Round(Rs!actup, 2)
li.SubItems(3) = Rs!salup
Rs.MoveNext
Loop
Exit Sub
hell:
If Err.Number = 3021 Then MsgBox "No Product Exits For this Supplier": lstv_ucost.ListItems.Clear
End Sub

Private Sub txtdiscR_Change()
txtNamount.Text = Val(txtGtotal.Text) - Val(txtdiscR.Text)

End Sub

Private Sub txtGtotal_Change()
txtNamount.Text = txtGtotal.Text
End Sub

Private Sub txtkg_Change()
On Error GoTo hell
If txtkg.Text > 0 Then
txtTqty.Text = Val(DubKg) * Val(txtkg.Text)
End If
Exit Sub
hell:
If Err.Number = 13 Then txtTqty.Text = ""
End Sub
Private Sub txtbox_Change()
On Error GoTo hell
If txtbox.Text > 0 Then
txtTqty.Text = Val(dubbox) * Val(txtbox.Text)
End If
Exit Sub
hell:
If Err.Number = 13 Then txtTqty.Text = ""
End Sub
Private Sub txtcarton_Change()
On Error GoTo hell
If txtcarton.Text > 0 Then
txtTqty.Text = Val(dubcarton) * Val(txtcarton.Text)
End If
Exit Sub
hell:
If Err.Number = 13 Then txtTqty.Text = ""
End Sub

Private Sub txtpieces_Change()
If txtkg.Locked = False Then
txtTqty.Text = Val(DubKg) * Val(txtkg.Text) + Val(txtpieces.Text)
End If
If txtbox.Locked = False Then
txtTqty.Text = Val(dubbox) * Val(txtbox.Text) + Val(txtpieces.Text)
End If
If txtcarton.Locked = False Then
txtTqty.Text = Val(dubcarton) * Val(txtcarton.Text) + Val(txtpieces.Text)
End If
If txtstrip.Locked = False Then
txtTqty.Text = Val(dubstrip) * Val(txtstrip.Text) + Val(txtpieces.Text)
End If
End Sub

Private Sub txtstrip_Change()
On Error GoTo hell
If txtstrip.Text > 0 Then
txtTqty.Text = Val(dubstrip) * Val(txtstrip.Text)
End If
Exit Sub
hell:
If Err.Number = 13 Then txtTqty.Text = ""
End Sub
Private Sub txtTqty_Change()
txtTamount.Text = Val(txtpup.Text) * Val(txtTqty.Text)
End Sub
Private Sub ClsTextBox()
txtkg.Text = ""
txtbox.Text = ""
txtcarton.Text = ""
txtstrip.Text = ""
txtpieces.Text = ""
End Sub
Private Sub GTotal_Amount()
Dim li As ListItem
Dim a As Integer
txtGtotal.Text = ""
For a = 1 To lstv_Purchase.ListItems.Count
Set li = lstv_Purchase.ListItems.item(a)
txtGtotal.Text = Val(txtGtotal.Text) + Val(li.SubItems(8))
Next a
lstv_Purchase.Refresh
End Sub

Private Sub productCoMbo()
Set Rs = New ADODB.Recordset
If Rs.State = 1 Then Rs.Close
Rs.Open "select product_id,product_name from qrypurchaseframe where sup_name='" & dc_supplier.Text & "'", cn, adOpenStatic, adLockOptimistic
Set dc_pname.RowSource = Rs
dc_pname.ListField = "product_name"
dc_pname.BoundColumn = "product_id"
End Sub
Private Sub ClearAll()
Call ClsTextBox
txtpup.Text = ""
dc_pname.Text = ""
End Sub
