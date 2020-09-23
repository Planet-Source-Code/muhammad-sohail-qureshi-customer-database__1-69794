VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvbutton.ocx"
Begin VB.Form frm_company 
   BackColor       =   &H00CDDAB1&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Company"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   11040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Customer.ctrlLiner ctrlLiner2 
      Height          =   30
      Left            =   -15
      TabIndex        =   24
      Top             =   585
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   53
   End
   Begin lvButton.lvButtons_H cmdnew 
      Height          =   690
      Left            =   525
      TabIndex        =   16
      Top             =   6300
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1217
      Caption         =   "New"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   4
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
      Image           =   "frm_company.frx":0000
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00CDDAB1&
      Height          =   5040
      Left            =   60
      TabIndex        =   15
      Top             =   690
      Width           =   5025
      Begin lvButton.lvButtons_H cmdfind 
         Height          =   330
         Left            =   300
         TabIndex        =   23
         Top             =   300
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
      Begin VB.PictureBox Picture1 
         Height          =   285
         Left            =   270
         ScaleHeight     =   225
         ScaleWidth      =   210
         TabIndex        =   22
         Top             =   975
         Width           =   270
      End
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
         Left            =   930
         TabIndex        =   21
         Top             =   315
         Width           =   3225
      End
      Begin MSComctlLib.ListView lstv_companyV 
         Height          =   4155
         Left            =   210
         TabIndex        =   20
         Top             =   705
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   7329
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
            Text            =   "Com id"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Company Name"
            Object.Width           =   6174
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00BBCD96&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   30
      TabIndex        =   12
      Top             =   0
      Width           =   4995
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H0095B165&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   -765
         TabIndex        =   13
         Top             =   60
         Width           =   5655
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "View Add Companies"
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
            Left            =   870
            TabIndex        =   14
            Top             =   45
            Width           =   3915
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDDAB1&
      Height          =   5055
      Left            =   5220
      TabIndex        =   4
      Top             =   690
      Width           =   5655
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1380
         TabIndex        =   6
         Top             =   375
         Width           =   1335
      End
      Begin VB.TextBox txtname 
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
         TabIndex        =   5
         Top             =   1440
         Width           =   4215
      End
      Begin lvButton.lvButtons_H cmdcancel 
         Height          =   585
         Left            =   2520
         TabIndex        =   7
         Top             =   2235
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
         Image           =   "frm_company.frx":077A
         ImgSize         =   24
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdsave 
         Height          =   585
         Left            =   885
         TabIndex        =   8
         Top             =   2220
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
         Image           =   "frm_company.frx":0EF4
         ImgSize         =   24
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdEsave 
         Height          =   585
         Left            =   900
         TabIndex        =   9
         Top             =   2235
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
         Image           =   "frm_company.frx":166E
         ImgSize         =   24
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H lvButtons_H2 
         Height          =   735
         Left            =   1650
         TabIndex        =   19
         Top             =   3030
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   1296
         Caption         =   "Refresh"
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   4
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
         Image           =   "frm_company.frx":1DE8
         ImgSize         =   24
         cBack           =   -2147483633
      End
      Begin VB.Label Label3 
         BackColor       =   &H00F8968B&
         BackStyle       =   0  'Transparent
         Caption         =   "Company ID"
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
         Left            =   165
         TabIndex        =   11
         Top             =   435
         Width           =   1320
      End
      Begin VB.Label Label1 
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
         Left            =   255
         TabIndex        =   10
         Top             =   1110
         Width           =   1680
      End
   End
   Begin Customer.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   30
      TabIndex        =   3
      Top             =   5865
      Width           =   10905
      _ExtentX        =   19235
      _ExtentY        =   53
   End
   Begin VB.Frame Frame7 
      Appearance      =   0  'Flat
      BackColor       =   &H00BBCD96&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   5070
      TabIndex        =   0
      Top             =   15
      Width           =   5925
      Begin VB.Frame Frame8 
         Appearance      =   0  'Flat
         BackColor       =   &H0095B165&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   -765
         TabIndex        =   1
         Top             =   60
         Width           =   6615
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Add New Company"
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
            Left            =   870
            TabIndex        =   2
            Top             =   45
            Width           =   3915
         End
      End
   End
   Begin lvButton.lvButtons_H cmdedit 
      Height          =   690
      Left            =   2107
      TabIndex        =   17
      Top             =   6300
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1217
      Caption         =   "Edit"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   4
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
      Image           =   "frm_company.frx":2562
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   690
      Left            =   3690
      TabIndex        =   18
      Top             =   6300
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1217
      Caption         =   "Delete"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   4
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
      Image           =   "frm_company.frx":2CDC
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdexit 
      Height          =   885
      Left            =   8865
      TabIndex        =   25
      Top             =   6195
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   1561
      Caption         =   "Exit"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   4
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
      Image           =   "frm_company.frx":3456
      ImgSize         =   32
      cBack           =   -2147483633
   End
End
Attribute VB_Name = "frm_company"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset
Dim Sql As String

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call SetListViewColorShort(lstv_companyV, Picture1, vbWhite, vbGray)
Call FrmCnt(frm_company)
Call find_Lock
Call cmPYId
Call List_comPany
End Sub
Private Sub cmdsave_Click()
Set Rs = New ADODB.Recordset
If Rs.State = 1 Then Rs.Close
Sql = "select * from company"
Rs.Open Sql, cn, adOpenDynamic, adLockOptimistic
Rs.AddNew
Rs!com_id = txtid.Text
Rs!company_name = txtname.Text
Rs.Update
End Sub
Private Sub cmPYId()
Set Rs = New ADODB.Recordset
If Rs.State = 1 Then Rs.Close
Sql = "select max(com_id)+1 as id from company"
Rs.Open Sql, cn, adOpenDynamic, adLockOptimistic
If IsNull(Rs!Id) Then
txtid.Text = 1
Else
txtid.Text = Rs!Id
End If
End Sub
Private Sub find_Lock()
txtfind.BackColor = &HC0FFFF
txtfind.Locked = True
End Sub
Private Sub List_comPany()
Dim li As ListItem
Set Rs = New ADODB.Recordset
If Rs.State = 1 Then Rs.Close
Rs.Open "select * from company", cn, adOpenDynamic, adLockReadOnly
lstv_companyV.ListItems.Clear
Do While Not Rs.EOF
Set li = lstv_companyV.ListItems.Add
li.Text = Rs!com_id
li.SubItems(1) = Rs!company_name
Rs.MoveNext
Loop
End Sub
