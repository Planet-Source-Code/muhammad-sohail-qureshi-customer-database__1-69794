VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_newbackAC 
   BackColor       =   &H00CDDAB1&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Bank Account"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtamtdeposit 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   1410
      TabIndex        =   3
      Top             =   2670
      Width           =   1305
   End
   Begin VB.TextBox txtacHname 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   3465
      TabIndex        =   2
      Top             =   2085
      Width           =   1950
   End
   Begin VB.TextBox txtacno 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   1185
      TabIndex        =   1
      Top             =   2085
      Width           =   1305
   End
   Begin VB.TextBox txtbankname 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   1185
      TabIndex        =   0
      Top             =   1575
      Width           =   4245
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00BBCD96&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   7230
      TabIndex        =   4
      Top             =   795
      Width           =   7230
      Begin VB.PictureBox Picture3 
         BackColor       =   &H009AB564&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   60
         ScaleHeight     =   285
         ScaleWidth      =   7380
         TabIndex        =   5
         Top             =   75
         Width           =   7380
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "New Bank Account"
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
            Left            =   45
            TabIndex        =   6
            Top             =   15
            Width           =   2040
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1965
      Top             =   3135
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_newbackAC.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_newbackAC.frx":077A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_newbackAC.frx":0EF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_newbackAC.frx":166E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_newbackAC.frx":1DE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_newbackAC.frx":2562
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_newbackAC.frx":2CDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_newbackAC.frx":3456
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_newbackAC.frx":3BD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_newbackAC.frx":434A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_newbackAC.frx":4AC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_newbackAC.frx":523E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker dp_bank 
      Height          =   345
      Left            =   3450
      TabIndex        =   12
      Top             =   2580
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   609
      _Version        =   393216
      CalendarBackColor=   16570065
      CalendarForeColor=   -2147483647
      CalendarTitleBackColor=   -2147483639
      CalendarTitleForeColor=   -2147483641
      CalendarTrailingForeColor=   -2147483635
      Format          =   20643841
      CurrentDate     =   39231
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   405
      Left            =   2865
      TabIndex        =   11
      Top             =   2700
      Width           =   720
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount Deposit"
      Height          =   255
      Left            =   225
      TabIndex        =   10
      Top             =   2730
      Width           =   1230
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A/C Holder Name"
      Height          =   420
      Left            =   2490
      TabIndex        =   9
      Top             =   2085
      Width           =   990
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Account No"
      Height          =   285
      Left            =   270
      TabIndex        =   8
      Top             =   2145
      Width           =   990
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Name"
      Height          =   285
      Left            =   255
      TabIndex        =   7
      Top             =   1620
      Width           =   990
   End
End
Attribute VB_Name = "Frm_newbackAC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Call FrmCnt(Frm_newbackAC)
dp_bank.Value = Date
End Sub
