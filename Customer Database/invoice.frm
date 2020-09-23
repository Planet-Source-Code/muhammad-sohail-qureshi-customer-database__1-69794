VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvbutton.ocx"
Begin VB.Form frm_Invoice 
   BackColor       =   &H00CDDAB1&
   Caption         =   "Sales Invoice"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "invoice.frx":0000
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00CDDAB1&
      ForeColor       =   &H80000008&
      Height          =   8550
      Left            =   -15
      TabIndex        =   0
      Top             =   -105
      Width           =   11850
      Begin VB.ComboBox comProduct 
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
         Height          =   315
         Left            =   7500
         TabIndex        =   31
         Top             =   1155
         Width           =   3075
      End
      Begin VB.ComboBox comCus 
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
         Height          =   315
         Left            =   4080
         TabIndex        =   30
         Top             =   1140
         Width           =   3225
      End
      Begin VB.TextBox txtcustype 
         Alignment       =   2  'Center
         ForeColor       =   &H00000080&
         Height          =   405
         Left            =   10275
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   1770
         Visible         =   0   'False
         Width           =   870
      End
      Begin MSComctlLib.ListView inv_lstv 
         Height          =   2535
         Left            =   1080
         TabIndex        =   23
         Top             =   4860
         Width           =   9825
         _ExtentX        =   17330
         _ExtentY        =   4471
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Invoice No"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Product No"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Product Name"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Unit Price"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Quantity"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Amount"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Dis Rs."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Invoice Type"
            Object.Width           =   2117
         EndProperty
      End
      Begin VB.PictureBox Picture1 
         Height          =   195
         Left            =   975
         ScaleHeight     =   135
         ScaleWidth      =   165
         TabIndex        =   19
         Top             =   5010
         Width           =   225
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00BBCD96&
         ForeColor       =   &H00404000&
         Height          =   2550
         Left            =   1680
         TabIndex        =   2
         Top             =   1650
         Width           =   8265
         Begin VB.TextBox txtStockH 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2100
            TabIndex        =   32
            Top             =   1935
            Width           =   1155
         End
         Begin VB.TextBox txtdic 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5400
            TabIndex        =   29
            Top             =   1920
            Width           =   930
         End
         Begin VB.TextBox txtTamount 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5385
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   1320
            Width           =   1425
         End
         Begin VB.TextBox txtqty 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2085
            TabIndex        =   17
            Top             =   1380
            Width           =   1440
         End
         Begin VB.TextBox txtproname 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   2070
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   840
            Width           =   4740
         End
         Begin VB.TextBox txtuPrice 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5370
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   255
            Width           =   1380
         End
         Begin VB.TextBox txtprocode 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   315
            Width           =   1455
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Stock in Hand"
            ForeColor       =   &H80000007&
            Height          =   285
            Index           =   0
            Left            =   960
            TabIndex        =   33
            Top             =   1995
            Width           =   1095
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Discount Rs."
            ForeColor       =   &H80000007&
            Height          =   285
            Index           =   7
            Left            =   4305
            TabIndex        =   28
            Top             =   1980
            Width           =   990
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Product Name"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   8
            Left            =   990
            TabIndex        =   7
            Top             =   870
            Width           =   1395
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Product Code"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   6
            Left            =   990
            TabIndex        =   6
            Top             =   390
            Width           =   1395
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Unit Price"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   4
            Left            =   4590
            TabIndex        =   5
            Top             =   375
            Width           =   1395
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Amount"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   1
            Left            =   4185
            TabIndex        =   4
            Top             =   1365
            Width           =   1095
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Quantity"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   2
            Left            =   1335
            TabIndex        =   3
            Top             =   1395
            Width           =   660
         End
      End
      Begin VB.TextBox txtid 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   450
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   1110
         Width           =   1485
      End
      Begin lvButton.lvButtons_H cmdaddenty 
         Height          =   555
         Left            =   10815
         TabIndex        =   22
         Top             =   2610
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   979
         Caption         =   "Add Entery"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   16761024
         cFHover         =   16777215
         cBhover         =   11730989
         LockHover       =   3
         cGradient       =   8421631
         Mode            =   0
         Value           =   0   'False
         cBack           =   14832217
      End
      Begin lvButton.lvButtons_H cmdsave 
         Height          =   450
         Left            =   10830
         TabIndex        =   25
         Top             =   3255
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   794
         Caption         =   "Save"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   16761024
         cFHover         =   16777215
         cBhover         =   11730989
         LockHover       =   3
         cGradient       =   8421631
         Mode            =   0
         Value           =   0   'False
         cBack           =   14832217
      End
      Begin MSComCtl2.DTPicker dp_i 
         Height          =   345
         Left            =   2565
         TabIndex        =   26
         Top             =   1125
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   609
         _Version        =   393216
         CalendarBackColor=   16570065
         CalendarForeColor=   -2147483647
         CalendarTitleBackColor=   -2147483639
         CalendarTitleForeColor=   -2147483641
         CalendarTrailingForeColor=   -2147483635
         Format          =   20578305
         CurrentDate     =   39231
      End
      Begin lvButton.lvButtons_H cmdcancel 
         Height          =   375
         Left            =   7335
         TabIndex        =   34
         Top             =   480
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   661
         Caption         =   "Cancel"
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
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   585
         Top             =   2250
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
               Picture         =   "invoice.frx":000C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "invoice.frx":199E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "invoice.frx":3330
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "invoice.frx":3AAA
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "invoice.frx":4224
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "invoice.frx":499E
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "invoice.frx":5118
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "invoice.frx":5892
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "invoice.frx":7224
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "invoice.frx":8BB6
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "invoice.frx":A548
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "invoice.frx":BEDA
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "invoice.frx":D86C
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "invoice.frx":F1FE
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "invoice.frx":10B92
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "invoice.frx":1186E
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "invoice.frx":1214E
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "invoice.frx":12E2A
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "invoice.frx":13B06
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "invoice.frx":147E2
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "invoice.frx":154BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "invoice.frx":1619A
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "invoice.frx":16A76
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Sales Detail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   1095
         TabIndex        =   27
         Top             =   4620
         Width           =   3015
      End
      Begin VB.Shape shpBar 
         BackColor       =   &H00E77A4B&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   240
         Left            =   1065
         Top             =   4590
         Width           =   9840
      End
      Begin VB.Label Label9 
         BackColor       =   &H00B6EBEA&
         Height          =   345
         Left            =   450
         TabIndex        =   21
         Top             =   1185
         Width           =   1545
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8415
         TabIndex        =   13
         Top             =   915
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   5025
         TabIndex        =   12
         Top             =   900
         Width           =   1020
      End
      Begin VB.Label Label5 
         BackColor       =   &H009AB564&
         Height          =   2310
         Index           =   0
         Left            =   1800
         TabIndex        =   11
         Top             =   2040
         Width           =   8265
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   2925
         TabIndex        =   10
         Top             =   870
         Width           =   525
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   465
         TabIndex        =   9
         Top             =   825
         Width           =   1605
      End
      Begin VB.Label Label5 
         BackColor       =   &H009AB564&
         Caption         =   "Label5"
         ForeColor       =   &H80000005&
         Height          =   2445
         Index           =   1
         Left            =   1410
         TabIndex        =   8
         Top             =   4920
         Width           =   9600
      End
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   495
      Left            =   5340
      TabIndex        =   20
      Top             =   4005
      Width           =   1215
   End
End
Attribute VB_Name = "frm_Invoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset
Private Sub Save()
Dim li As ListItem
Dim i As Integer
If inv_lstv.ListItems.Count = 0 Then Exit Sub
For i = 1 To inv_lstv.ListItems.Count
Set li = inv_lstv.ListItems.Item(i)
If Rs.State = 1 Then Rs.Close
Rs.Open "select * from invoice_detail", cn, adOpenDynamic, adLockOptimistic
Rs.AddNew
If li.SubItems(1) <> "" Then
Rs!inv_id = li.Text
Rs!item_id = li.SubItems(1)
Rs!item_name = li.SubItems(2)
Rs!unit_price = li.SubItems(3)
Rs!quantity = li.SubItems(4)
Rs!amount = li.SubItems(5)
Rs!dis = li.SubItems(6)
Rs!inv_type = li.SubItems(7)
End If
Rs.Update
Next
MsgBox "Save Successfully", vbInformation + vbOKOnly, "Soft Vision"
If vbYes Then
cmdsave.Enabled = False
cmdaddenty.Enabled = False
cmdcancel.Enabled = False
End If
If Rs.State = 1 Then Rs.Close
Rs.Open "select * from invoice", cn, adOpenDynamic, adLockOptimistic
Rs.AddNew
Rs!invoice_id = txtid.Text
Rs!cus_name = dc_cusname.Text
Rs!date_i = dp_i.Value
Rs.Update
End Sub
Private Sub dc_cusname_Click(Area As Integer)
If dc_cusname.Locked = True Then
MsgBox "You can't select another Customer right now..", vbInformation, ""
Exit Sub
End If
End Sub

Private Sub Form_Load()
Call Customer_Combo
'____________________

Frame4.Enabled = False

cmdaddenty.Enabled = False
txtid.Enabled = False
dp_i.Enabled = False
cmdsave.Enabled = False
'________________________

dp_i.Value = Date
Call SetListViewColor(inv_lstv, Picture1, vbWhite, vbGray)
End Sub
Private Sub Form_Unload(Cancel As Integer)
Call Mdipicu
Set frm_cshsale = Nothing
End Sub
Private Sub Serial()
If Rs.State = 1 Then Rs.Close
Rs.Open "select max(invoice_id)+1 as id from invoice", cn, adOpenDynamic, adLockOptimistic
If IsNull(Rs!Id) Then
txtid.Text = 1
Else
txtid.Text = Rs!Id
End If
End Sub
Private Sub inv_lstv_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
Dim li As ListItem
Set li = inv_lstv.SelectedItem
MsgBox "Do you want to Delete it!", vbInformation + vbYesNo, "Soft Vision"
If vbYes Then
inv_lstv.ListItems.Remove li.Index
If vbNo Then Exit Sub
End If
End If
End Sub

Private Sub txtdic_Change()
txtTamount.Text = (Val(txtuPrice.Text) * Val(txtqty.Text) - Val(txtdic.Text))
End Sub
Private Sub txtqty_Change()
txtTamount.Text = (Val(txtuPrice.Text) * Val(txtqty.Text))
End Sub
Private Sub Customer_Combo()
Set Rs = New ADODB.Recordset
If Rs.State = 1 Then Rs.Close
Rs.Open "select name from customer", cn, adOpenDynamic, adLockOptimistic
While Not Rs.EOF
comCus.AddItem Rs!Name
Rs.MoveNext
Wend
End Sub
Private Sub Product_combo()
Set Rs = New ADODB.Recordset
If Rs.State = 1 Then Rs.Close

End Sub
