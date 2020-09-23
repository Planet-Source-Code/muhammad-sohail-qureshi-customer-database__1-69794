VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvbutton.ocx"
Begin VB.Form frm_opStock 
   BackColor       =   &H00CDDAB1&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Opening Stock"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   10695
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   10695
   Begin VB.Frame Frame2 
      BackColor       =   &H00CDDAB1&
      Height          =   2115
      Left            =   75
      TabIndex        =   22
      Top             =   5430
      Width           =   10515
      Begin VB.Frame Frame4 
         BackColor       =   &H00CDDAB1&
         Height          =   1890
         Left            =   3195
         TabIndex        =   31
         Top             =   120
         Width           =   2760
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Double Click for Edit"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   465
            TabIndex        =   32
            Top             =   390
            Width           =   1785
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00CDDAB1&
         Height          =   1890
         Left            =   6345
         TabIndex        =   24
         Top             =   120
         Width           =   4080
         Begin lvButton.lvButtons_H lvButtons_H1 
            Height          =   465
            Left            =   225
            TabIndex        =   25
            Top             =   315
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   820
            Caption         =   "Open"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFore           =   16711680
            cFHover         =   16711680
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin lvButton.lvButtons_H lvButtons_H2 
            Height          =   465
            Left            =   1477
            TabIndex        =   26
            Top             =   315
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   820
            Caption         =   "Save"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFore           =   16711680
            cFHover         =   16711680
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin lvButton.lvButtons_H lvButtons_H3 
            Height          =   465
            Left            =   2730
            TabIndex        =   27
            Top             =   315
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   820
            Caption         =   "Exit"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFore           =   16711680
            cFHover         =   16711680
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin lvButton.lvButtons_H lvButtons_H4 
            Height          =   465
            Left            =   225
            TabIndex        =   28
            Top             =   1170
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   820
            Caption         =   "New"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFore           =   16711680
            cFHover         =   16711680
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin lvButton.lvButtons_H lvButtons_H5 
            Height          =   465
            Left            =   1477
            TabIndex        =   29
            Top             =   1170
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   820
            Caption         =   "Edit"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFore           =   16711680
            cFHover         =   16711680
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin lvButton.lvButtons_H lvButtons_H6 
            Height          =   465
            Left            =   2730
            TabIndex        =   30
            Top             =   1170
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   820
            Caption         =   "Delete"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFore           =   16711680
            cFHover         =   16711680
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1725
         Left            =   135
         TabIndex        =   23
         Top             =   225
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   3043
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
            Text            =   "Opid"
            Object.Width           =   1676
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   2822
         EndProperty
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00BBCD96&
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   0
      TabIndex        =   19
      Top             =   75
      Width           =   10650
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         BackColor       =   &H0095B165&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   90
         TabIndex        =   20
         Top             =   45
         Width           =   10500
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Opening Stock"
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
            TabIndex        =   21
            Top             =   30
            Width           =   2100
         End
      End
   End
   Begin Customer.ctrlLiner ctrlLiner2 
      Height          =   30
      Left            =   -30
      TabIndex        =   16
      Top             =   4740
      Width           =   10680
      _ExtentX        =   18838
      _ExtentY        =   53
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDDAB1&
      Height          =   3600
      Left            =   75
      TabIndex        =   5
      Top             =   1065
      Width           =   10500
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
         Left            =   180
         TabIndex        =   11
         Top             =   480
         Width           =   990
      End
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
         Left            =   1320
         TabIndex        =   10
         Top             =   510
         Width           =   4935
      End
      Begin VB.TextBox Text3 
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
         Left            =   7380
         TabIndex        =   9
         Top             =   510
         Width           =   810
      End
      Begin VB.TextBox Text4 
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
         Left            =   6390
         TabIndex        =   8
         Top             =   510
         Width           =   810
      End
      Begin MSComctlLib.ListView lstv_op 
         Height          =   2520
         Left            =   300
         TabIndex        =   6
         Top             =   960
         Width           =   9930
         _ExtentX        =   17515
         _ExtentY        =   4445
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
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ProID"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Pro Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Qty"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Rate"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Amount"
            Object.Width           =   3175
         EndProperty
      End
      Begin lvButton.lvButtons_H cmdadd 
         Height          =   405
         Left            =   8475
         TabIndex        =   7
         Top             =   495
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   714
         Caption         =   "Add"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         BackColor       =   &H00F8968B&
         BackStyle       =   0  'Transparent
         Caption         =   "Pro ID*"
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
         Left            =   255
         TabIndex        =   15
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label4 
         BackColor       =   &H00F8968B&
         BackStyle       =   0  'Transparent
         Caption         =   "Product Name"
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
         Left            =   1365
         TabIndex        =   14
         Top             =   240
         Width           =   1530
      End
      Begin VB.Label Label5 
         BackColor       =   &H00F8968B&
         BackStyle       =   0  'Transparent
         Caption         =   "Rate*"
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
         Left            =   7515
         TabIndex        =   13
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label6 
         BackColor       =   &H00F8968B&
         BackStyle       =   0  'Transparent
         Caption         =   "Qty*"
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
         Left            =   6585
         TabIndex        =   12
         Top             =   195
         Width           =   585
      End
   End
   Begin Customer.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   -15
      TabIndex        =   4
      Top             =   1035
      Width           =   10665
      _ExtentX        =   18812
      _ExtentY        =   53
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   600
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   20774913
      CurrentDate     =   39424
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
      Left            =   2190
      TabIndex        =   0
      Top             =   615
      Width           =   990
   End
   Begin VB.Label Label7 
      BackColor       =   &H00F8968B&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
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
      Left            =   6345
      TabIndex        =   18
      Top             =   4995
      Width           =   1410
   End
   Begin VB.Label lbltotal 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7710
      TabIndex        =   17
      Top             =   4920
      Width           =   1500
   End
   Begin VB.Label Label1 
      BackColor       =   &H00F8968B&
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
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
      Left            =   3495
      TabIndex        =   2
      Top             =   675
      Width           =   570
   End
   Begin VB.Label Label3 
      BackColor       =   &H00F8968B&
      BackStyle       =   0  'Transparent
      Caption         =   "Open ID*"
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
      Left            =   1215
      TabIndex        =   1
      Top             =   660
      Width           =   945
   End
End
Attribute VB_Name = "frm_opStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call FrmCnt(frm_opStock)
'LoadForm frm_opStock
'MdItoolBar
End Sub

Private Sub Form_Unload(Cancel As Integer)
MditoolbarV
End Sub
