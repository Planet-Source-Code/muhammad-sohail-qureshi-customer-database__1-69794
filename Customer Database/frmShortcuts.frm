VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvbutton.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmShortcuts 
   BackColor       =   &H00CDDAB1&
   Caption         =   "Shortcuts"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7350
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmShortcuts.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   7350
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame13 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   2265
      TabIndex        =   13
      Top             =   5715
      Width           =   9525
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   825
         Left            =   15
         TabIndex        =   14
         Top             =   -15
         Width           =   9585
         Begin VB.Image Image5 
            Height          =   240
            Left            =   6840
            Picture         =   "frmShortcuts.frx":08CA
            Top             =   375
            Width           =   240
         End
         Begin VB.Label lblsupplier 
            BackColor       =   &H80000012&
            Caption         =   "Label4"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   360
            Left            =   3900
            TabIndex        =   23
            Top             =   345
            Width           =   3495
         End
         Begin VB.Image Image4 
            Height          =   240
            Left            =   3150
            Picture         =   "frmShortcuts.frx":0C54
            Top             =   345
            Width           =   240
         End
         Begin VB.Label lblcustomer 
            BackColor       =   &H80000012&
            Caption         =   "Label4"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   360
            Left            =   210
            TabIndex        =   15
            Top             =   345
            Width           =   3495
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CDDAB1&
      BorderStyle     =   0  'None
      Height          =   840
      Left            =   -15
      TabIndex        =   12
      Top             =   5625
      Width           =   2310
      Begin VB.Image Image3 
         Height          =   825
         Left            =   30
         Picture         =   "frmShortcuts.frx":0FDE
         Top             =   45
         Width           =   2250
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   20
      Left            =   4035
      Top             =   6855
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   2220
      Left            =   -15
      ScaleHeight     =   2220
      ScaleWidth      =   12150
      TabIndex        =   2
      Top             =   -15
      Width           =   12150
      Begin VB.Timer Timer1 
         Interval        =   10
         Left            =   5040
         Top             =   105
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1605
         Left            =   15
         ScaleHeight     =   1605
         ScaleWidth      =   3720
         TabIndex        =   8
         Top             =   30
         Width           =   3720
         Begin lvButton.lvButtons_H lvButtons_H1 
            Height          =   495
            Left            =   2085
            TabIndex        =   22
            Top             =   1020
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   873
            Caption         =   "About Developer"
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
         Begin VB.Frame Frame6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   1890
            TabIndex        =   21
            Top             =   0
            Width           =   1860
         End
         Begin VB.Frame Frame5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            ForeColor       =   &H80000008&
            Height          =   1560
            Left            =   1860
            TabIndex        =   20
            Top             =   0
            Width           =   195
         End
         Begin VB.Frame Frame3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   2055
            TabIndex        =   17
            Top             =   1200
            Width           =   1860
         End
         Begin VB.Frame Frame4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            ForeColor       =   &H80000008&
            Height          =   1590
            Left            =   3300
            TabIndex        =   18
            Top             =   -15
            Width           =   525
         End
         Begin SHDocVwCtl.WebBrowser wb2 
            Height          =   1875
            Left            =   1875
            TabIndex        =   16
            Top             =   -120
            Width           =   2430
            ExtentX         =   4286
            ExtentY         =   3307
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   1
            AutoArrange     =   0   'False
            NoClientEdge    =   0   'False
            AlignLeft       =   0   'False
            NoWebView       =   0   'False
            HideFileNames   =   0   'False
            SingleClick     =   0   'False
            SingleSelection =   0   'False
            NoFolders       =   0   'False
            Transparent     =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   "http:///"
         End
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1560
            Left            =   60
            Picture         =   "frmShortcuts.frx":23F6
            ScaleHeight     =   1560
            ScaleWidth      =   2205
            TabIndex        =   9
            Top             =   0
            Width           =   2205
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            Height          =   375
            Left            =   1725
            TabIndex        =   19
            Top             =   1170
            Width           =   1650
         End
      End
      Begin VB.PictureBox Picture8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   9360
         ScaleHeight     =   285
         ScaleWidth      =   2550
         TabIndex        =   7
         Top             =   0
         Width           =   2550
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2100
         Left            =   11415
         ScaleHeight     =   2100
         ScaleWidth      =   780
         TabIndex        =   6
         Top             =   -105
         Width           =   780
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   9360
         ScaleHeight     =   240
         ScaleWidth      =   3090
         TabIndex        =   5
         Top             =   1560
         Width           =   3090
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1845
         Left            =   9330
         ScaleHeight     =   1845
         ScaleWidth      =   375
         TabIndex        =   4
         Top             =   -120
         Width           =   375
      End
      Begin SHDocVwCtl.WebBrowser wb 
         CausesValidation=   0   'False
         DragMode        =   1  'Automatic
         Height          =   1935
         Left            =   9615
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   -270
         Width           =   2400
         ExtentX         =   4233
         ExtentY         =   3413
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   0
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin VB.Image nScroll 
         Height          =   1350
         Left            =   -5430
         Picture         =   "frmShortcuts.frx":304B
         Top             =   135
         Width           =   8205
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Ph No : 0606-309345  Cell No : 0334-6943886"
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1875
         Width           =   3495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Developed By : Muhammad Sohail Qureshi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   30
         TabIndex        =   10
         Top             =   1635
         Width           =   3615
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3240
      Left            =   -15
      ScaleHeight     =   3240
      ScaleWidth      =   11985
      TabIndex        =   0
      Top             =   2190
      Width           =   11985
      Begin MSComctlLib.ListView lvMenu 
         Height          =   2805
         Left            =   30
         TabIndex        =   1
         Top             =   300
         Width           =   11805
         _ExtentX        =   20823
         _ExtentY        =   4948
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         OLEDragMode     =   1
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         MousePointer    =   99
         MouseIcon       =   "frmShortcuts.frx":61D0
         OLEDragMode     =   1
         NumItems        =   0
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2655
      Top             =   6615
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   26
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":6332
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":7CC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":89A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":A332
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":BCC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":D656
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":EFE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":FCC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":1099C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":11676
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":12352
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":1302E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":1390A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":145E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":152C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":15F9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":16882
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":1755E
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":17E3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":18B16
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":1A4AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":1BE3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":1C71A
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":229B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":28C4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":34CA0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image2 
      Height          =   165
      Left            =   9105
      Picture         =   "frmShortcuts.frx":414FA
      Top             =   5460
      Width           =   9300
   End
   Begin VB.Image Image1 
      Height          =   165
      Left            =   -165
      Picture         =   "frmShortcuts.frx":42201
      Top             =   5460
      Width           =   9300
   End
End
Attribute VB_Name = "frmShortcuts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Rs As ADODB.Recordset
Private Sub Form_Load()
Call cuStomer_Count: Call supplier_count
wb2.Navigate App.Path & "\copy2.gif"



nScroll.Left = Me.Width - nScroll.Width ' For Moving Label
wb.Navigate App.Path & "\Allah Name.gif"
    With lvMenu
        Set .SmallIcons = ImageList1
        Set .Icons = ImageList1
        'For Sales
        '.ListItems.Add , "frmCustomer", "Manage Customer", 1, 1
        .ListItems.Add , "frm_customer", "Customers Setup", 2, 2
        .ListItems.Add , "frmAccCustomer", "Customer Accounts", 18, 18
        .ListItems.Add , "frmCustomerWB", "Customers with Balance", 22, 22
        
      '  .ListItems.Add , "frmSalesman", "Manage Salesman", 3, 3
        .ListItems.Add , "frm_area", "Area Setup", 7, 7
        
        .ListItems.Add , "frmPDCManager", "PDC Manager", 12, 12
        .ListItems.Add , "frmDueChecks", "Display Due Checks", 13, 13
        
        'For Inventory
        .ListItems.Add , "frmSupplier", "Manage Suppliers", 4, 4
    
        .ListItems.Add , "frm_purchaseV", "Purchases", 5, 5
        .ListItems.Add , "frm_pdetailV", "New Product ", 6, 6
        
        .ListItems.Add , "frm_generic", "Product Generic Setup", 9, 9
        .ListItems.Add , "frm_opStock", "Opening Stock", 8, 8
        
        'For Transaction
        '.ListItems.Add , "frmLoading", "Van Loading", 10, 10
        .ListItems.Add , "frmInvoice", "Sales Invoice", 14, 14
        '.ListItems.Add , "frmVanCollection", "Van Collection", 15, 15
        '.ListItems.Add , "frmVanInventory", "Van Inventory", 11, 11
        .ListItems.Add , "frmcompanyV", "Supplier's Payable", 19, 19
        
        '.ListItems.Add , "frmSelectZipCode", "Manage Zip Codes", 20, 20
        .ListItems.Add , "frm_company", "Company Setup", 21, 21
        .ListItems.Add , "frmUserRec", "User Records", 17, 17
        .ListItems.Add , "frmBusinessInfo", "Business Information", 16, 16
        .ListItems.Add , "Sales Return", "Sales Return", 23, 23
        .ListItems.Add , "Purchase Return", "Purchase Return", 24, 24
        .ListItems.Add , "frm_PType", "Product Type Setup", 25, 25
        .ListItems.Add , "Exit", "Exit From Database", 26, 26
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Beep: Cancel = 1
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    lvMenu.Width = ScaleWidth
    lvMenu.Height = ScaleHeight
End Sub

Private Sub lvButtons_H1_Click()
Form1.Show
End Sub

Private Sub lvMenu_DblClick()
    Select Case lvMenu.SelectedItem.Key
        'For Sales
        'Case "frmCustomer": LoadForm frm_customer
        Case "frm_customer":  frm_customer.Show
        'Case "frmAccCustomer": LoadForm frmAccCustomer
        'Case "frmCustomerWB": LoadForm frmCustomerWB
            
        'Case "frmSalesman": LoadForm frmSalesman
        Case "frm_area": frm_area.Show
        
       ' Case "frmPDCManager": LoadForm frmPDCManager
        'Case "frmDueChecks": LoadForm frmDueChecks
    
        'For Inventory
        Case "frmSupplier": LoadForm frm_supView
            
        Case "frm_purchaseV": LoadForm frm_purchaseV
        Case "frm_pdetailV": LoadForm frm_ProductV
        
        Case "frm_generic": frm_generic.Show
        Case "frm_opStock":  frm_opStock.Show
        
        'For Transaction
        'Case "frmInvoice": LoadForm frmInvoice
        'Case "frmLoading": LoadForm frmLoading
        'Case "frmVanInventory": LoadForm frmVanInventory
        
        'Case "frmVanCollection": LoadForm frmVanCollection
        'Case "frmVanRemmitance": LoadForm frmVanRemmitance
        
        'Case "frmUserRec"
         '   If CurrUser.USER_ISADMIN = False Then
          '      MsgBox "Only admin users can access this record.", vbCritical, "Access Denied"
           ' Else
            '    frmUserRec.Show vbModal
           ' End If
        'Case "frmBusinessInfo": frmBusinessInfo.Show vbModal
        
        'Case "frmSelectZipCode": frmSelectZipCode.OPEN_COMMAND = 1: frmSelectZipCode.Show vbModal
        'Case "frmSelectBank": frmSelectBank.OPEN_COMMAND = 1: frmSelectBank.Show vbModal
        Case "frm_company":  frm_company.Show
        Case "frm_PType":  frm_PType.Show
         Case "Exit": End
    End Select

End Sub

Private Sub lvMenu_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call lvMenu_DblClick
End If
End Sub

Private Sub Timer1_Timer()
'Timer for help in moving image name nScroll
nScroll.Left = nScroll.Left - 10
 If nScroll.Left < nScroll.Width * -1 Then
  nScroll.Left = Me.Width
 End If
End Sub

Private Sub Timer2_Timer()
Frame1.Left = Frame1.Left - 10
If Frame1.Left < Frame1.Width * -1 Then
Frame1.Left = Me.Width
End If
End Sub
Private Sub cuStomer_Count()
Set Rs = New ADODB.Recordset
If Rs.State = 1 Then Rs.Close
Rs.Open "select * from customer ", cn, adOpenStatic, adLockOptimistic
lblcustomer.Caption = "You Have" & " " & Rs.RecordCount & "  " & "Customers"
End Sub
Private Sub supplier_count()
Set Rs = New ADODB.Recordset
If Rs.State = 1 Then Rs.Close
Rs.Open "select * from supplier", cn, adOpenStatic, adLockOptimistic
lblsupplier.Caption = "You Have" & " " & Rs.RecordCount & " " & "Suppliers"
End Sub

