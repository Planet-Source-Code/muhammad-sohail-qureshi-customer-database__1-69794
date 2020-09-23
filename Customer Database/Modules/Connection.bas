Attribute VB_Name = "Connection"
Option Explicit
Public cn As ADODB.Connection
'Private rs As ADODB.Recordset
Public Sub Conn()
Set cn = New ADODB.Connection
cn.CursorLocation = adUseClient
cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Customer.mdb;Persist Security Info=False"
cn.Open
End Sub
Public Sub LoadForm(ByRef srcForm As Form)
    srcForm.Show
    srcForm.WindowState = vbMaximized
    srcForm.SetFocus
End Sub
'Public Sub RsClose()
'Set rs = New ADODB.Recordset
'If rs.State = 1 Then rs.Close

'End Sub
