VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvbutton.ocx"
Begin VB.Form frm_generic 
   BackColor       =   &H00CDDAB1&
   Caption         =   "Form1"
   ClientHeight    =   6900
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9345
   LinkTopic       =   "Form1"
   ScaleHeight     =   6900
   ScaleWidth      =   9345
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      BackColor       =   &H00CDDAB1&
      Height          =   5025
      Left            =   135
      TabIndex        =   15
      Top             =   660
      Width           =   4110
      Begin VB.TextBox txtfind 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Height          =   330
         Left            =   795
         TabIndex        =   16
         Top             =   360
         Width           =   3120
      End
      Begin MSComctlLib.ListView lstv_area 
         Height          =   4200
         Left            =   105
         TabIndex        =   17
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Area Name"
            Object.Width           =   4762
         EndProperty
      End
      Begin lvButton.lvButtons_H cmdfind 
         Height          =   330
         Left            =   195
         TabIndex        =   18
         Top             =   345
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   582
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
      Left            =   4470
      TabIndex        =   12
      Top             =   0
      Width           =   4800
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H0095B165&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   60
         TabIndex        =   13
         Top             =   90
         Width           =   4650
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Add New Area"
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
            TabIndex        =   14
            Top             =   30
            Width           =   2100
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDDAB1&
      Height          =   5040
      Left            =   4305
      TabIndex        =   3
      Top             =   660
      Width           =   4950
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
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   480
         Width           =   1335
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
         Left            =   285
         TabIndex        =   4
         Top             =   1440
         Width           =   4275
      End
      Begin lvButton.lvButtons_H cmdcancel 
         Height          =   585
         Left            =   2895
         TabIndex        =   6
         Top             =   2850
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
         Image           =   "frm_generic.frx":0000
         ImgSize         =   24
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdsave 
         Height          =   585
         Left            =   1035
         TabIndex        =   7
         Top             =   2850
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
         Image           =   "frm_generic.frx":077A
         ImgSize         =   24
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdEsave 
         Height          =   585
         Left            =   1020
         TabIndex        =   8
         Top             =   2850
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
         Image           =   "frm_generic.frx":0EF4
         ImgSize         =   24
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdpprint 
         Height          =   690
         Left            =   1905
         TabIndex        =   9
         Top             =   3780
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   1217
         Caption         =   "Print"
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
         Image           =   "frm_generic.frx":166E
         ImgSize         =   32
         cBack           =   -2147483633
      End
      Begin VB.Label Label1 
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
         Left            =   120
         TabIndex        =   11
         Top             =   495
         Width           =   1050
      End
      Begin VB.Label Label3 
         BackColor       =   &H00F8968B&
         BackStyle       =   0  'Transparent
         Caption         =   "Area Name*"
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
         Left            =   315
         TabIndex        =   10
         Top             =   1125
         Width           =   1290
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00BBCD96&
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   4380
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         BackColor       =   &H0095B165&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   75
         TabIndex        =   1
         Top             =   45
         Width           =   4215
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "View Areas"
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
            TabIndex        =   2
            Top             =   30
            Width           =   2100
         End
      End
   End
   Begin Customer.ctrlLiner ctrlLiner2 
      Height          =   30
      Left            =   0
      TabIndex        =   19
      Top             =   615
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   53
   End
   Begin Customer.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   45
      TabIndex        =   20
      Top             =   5775
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   53
   End
   Begin lvButton.lvButtons_H cmdnew 
      Height          =   570
      Left            =   315
      TabIndex        =   21
      Top             =   6150
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
      Image           =   "frm_generic.frx":D6C0
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdedit 
      Height          =   570
      Left            =   1950
      TabIndex        =   22
      Top             =   6150
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
      Image           =   "frm_generic.frx":19712
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmddelete 
      Height          =   570
      Left            =   3585
      TabIndex        =   23
      Top             =   6135
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
      Image           =   "frm_generic.frx":19E8C
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdexit 
      Height          =   660
      Left            =   7140
      TabIndex        =   24
      Top             =   6060
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   1164
      Caption         =   "Exit"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
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
      Image           =   "frm_generic.frx":25EDE
      ImgSize         =   32
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdview 
      Height          =   570
      Left            =   5220
      TabIndex        =   25
      Top             =   6150
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
      Image           =   "frm_generic.frx":31F30
      ImgSize         =   24
      cBack           =   -2147483633
   End
End
Attribute VB_Name = "frm_generic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Call FrmCnt(frm_generic)
End Sub
