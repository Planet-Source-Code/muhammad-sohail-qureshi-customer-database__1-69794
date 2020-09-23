VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Customer Related Information"
   ClientHeight    =   5010
   ClientLeft      =   165
   ClientTop       =   255
   ClientWidth     =   5985
   Icon            =   "MDImain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picContainer 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   0
      ScaleHeight     =   900
      ScaleWidth      =   5985
      TabIndex        =   5
      Top             =   30
      Width           =   5985
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   780
         Left            =   60
         TabIndex        =   6
         Top             =   60
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1376
         ButtonWidth     =   1270
         ButtonHeight    =   1376
         Appearance      =   1
         Style           =   1
         ImageList       =   "MyImages2"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Shortcut"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Refresh"
               Key             =   "Refresh"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Help"
               Key             =   "Help"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   30
      Left            =   0
      ScaleHeight     =   30
      ScaleWidth      =   5985
      TabIndex        =   2
      Top             =   0
      Width           =   5985
      Begin VB.ListBox List1 
         Height          =   1620
         Left            =   4125
         TabIndex        =   3
         Top             =   1635
         Visible         =   0   'False
         Width           =   1380
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   765
         Top             =   2355
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   24
         ImageHeight     =   24
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDImain.frx":08CA
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDImain.frx":1044
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDImain.frx":17BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDImain.frx":1900
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label17 
         BackColor       =   &H8000000D&
         Height          =   1500
         Left            =   1200
         TabIndex        =   4
         Top             =   2355
         Visible         =   0   'False
         Width           =   2235
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   4695
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   556
      SimpleText      =   "Status Bar"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   758
            MinWidth        =   758
            Text            =   "Time"
            TextSave        =   "Time"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1499
            MinWidth        =   1499
            TextSave        =   "1:19 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "Date"
            TextSave        =   "Date"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "12/10/2007"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Width           =   8290
            MinWidth        =   8290
            Picture         =   "MDImain.frx":207A
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7497
            MinWidth        =   7497
            Text            =   "This Softwate is created by : Muhammad Sohail Qureshi"
            TextSave        =   "This Softwate is created by : Muhammad Sohail Qureshi"
         EndProperty
      EndProperty
      MousePointer    =   12
   End
   Begin VB.PictureBox Picture7 
      Align           =   1  'Align Top
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   5985
      TabIndex        =   1
      Top             =   930
      Width           =   5985
   End
   Begin MSComctlLib.ImageList itb32x32 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":6128
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":7ABA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":944C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":ADDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":C770
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":E102
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":FA94
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":11426
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":12DB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":1474C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":15428
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":15D08
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":169E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":176C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":1839C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":19078
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":19D54
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList MyImages2 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":1A630
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":1BFC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":1D954
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":1E630
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnu_dataentery 
      Caption         =   "Data Entery"
      Begin VB.Menu mnu_addcustomer 
         Caption         =   "Add Customers"
      End
      Begin VB.Menu dfd 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_sales 
         Caption         =   "Sales"
         Begin VB.Menu mnu_cashsales 
            Caption         =   "Cash Sales"
         End
         Begin VB.Menu fds 
            Caption         =   "-"
         End
         Begin VB.Menu mnu_creditsales 
            Caption         =   "Credit Sales"
         End
      End
      Begin VB.Menu fdf 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_purchase 
         Caption         =   "Purchases"
         Begin VB.Menu mnu_cashpurchase 
            Caption         =   "Cash Purchase"
         End
         Begin VB.Menu dfdsffd 
            Caption         =   "-"
         End
         Begin VB.Menu mnu_creditpur 
            Caption         =   "Credit Purchase"
         End
      End
      Begin VB.Menu gfdasf 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_cusreceving 
         Caption         =   "Customer Recevings"
      End
      Begin VB.Menu dfdsf 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_supplier 
         Caption         =   "Add Suppliers"
      End
      Begin VB.Menu jhgj 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_newproducts 
         Caption         =   "New Prodcusts"
      End
   End
   Begin VB.Menu mnu_Return 
      Caption         =   "Returns"
      Begin VB.Menu mnu_purReturn 
         Caption         =   "Purchase Return"
      End
      Begin VB.Menu gfdgfdg 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_salesreturn 
         Caption         =   "Sales Return"
      End
   End
   Begin VB.Menu mnu_bank 
      Caption         =   "Bank "
      Begin VB.Menu mnu_newbankac 
         Caption         =   "New Bank Account"
      End
      Begin VB.Menu fdslkfj 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_amtdeposit 
         Caption         =   "Amount Deposit"
      End
      Begin VB.Menu hgfhgh 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_amtIssue 
         Caption         =   "Amount Isuue"
      End
   End
   Begin VB.Menu mnu_view 
      Caption         =   "View"
      Begin VB.Menu mnu_customerview 
         Caption         =   "Customers"
      End
      Begin VB.Menu hjhgjg 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_suppliers 
         Caption         =   "Suppliers"
      End
      Begin VB.Menu jfklsdjf 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_purchasev 
         Caption         =   "Purchase"
      End
      Begin VB.Menu fdkjdsalkjf 
         Caption         =   "-"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub quit()
Dim ans As Variant
ans = MsgBox("Do you realy want to Exit", vbInformation + vbYesNo, "Exit")
If ans = vbYes Then
End
Else
Exit Sub
End If
End Sub
Private Sub MDIForm_Load()
Call MditoolbarV
Call Conn
Call Disk_Serial
If Not Label17.Caption = "KF112ZW" Then
MsgBox "Taking invald Process! All Rights Received to Softvision", vbOKOnly + vbInformation, "Soft Vision"
If cn.State = 1 Then cn.Close
End
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "Refresh"
frmShortcuts.lvMenu.Refresh
Case "Help"
MsgBox "For any query and help call freely Muhammad Sohail Qureshi at 0334-6943886 or 0304-6550770"
   End Select
End Sub
