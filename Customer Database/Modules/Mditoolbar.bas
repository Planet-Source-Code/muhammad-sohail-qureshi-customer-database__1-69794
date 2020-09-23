Attribute VB_Name = "Mditoolbar_action"
Option Explicit

Public Sub MdItoolBar()
MDIForm1.picContainer.Visible = False
End Sub
Public Sub MditoolbarV()
MDIForm1.picContainer.Visible = True
MDIForm1.Toolbar1.Visible = True
End Sub

Public Sub FrmCnt(fRm As Form)
fRm.Move (Screen.Width - fRm.Width) / 2, (Screen.Height - fRm.Height) / 2
End Sub

