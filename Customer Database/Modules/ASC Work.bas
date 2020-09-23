Attribute VB_Name = "Ascii"
Option Explicit
Public Sub TextValid(KeyAscii As Integer, charAccept As Integer)
If KeyAscii = vbKeyBack Then
Exit Sub
End If

If KeyAscii = vbKeyReturn Then
SendKeys "{tab}"
KeyAscii = 0
Exit Sub
End If


Select Case charAccept
Case 0 'any character
Exit Sub
Case 1 'only numbers
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 32 Then
Exit Sub
Else
KeyAscii = 0
Beep
MsgBox "Enter Only Numeric Value", vbExclamation, "Data Entery Error"
End If
Case 2 'only letters
If (KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 32 Then
Exit Sub
ElseIf (KeyAscii >= 97 And KeyAscii <= 122) Then
Exit Sub
Else
KeyAscii = 0
Beep
MsgBox "Enter Only Alphabet", vbExclamation, "Data Entery Error!"
End If
End Select
End Sub

Public Sub TextGotF(txt As TextBox) 'effect on focus
txt.BackColor = &H80C0FF
txt.ForeColor = vbBlue
txt.SelLength = 0
txt.SelLength = Len(txt.Text)
End Sub
Public Sub TextLostF(txt As TextBox) 'effect on lostfocus
txt.BackColor = vbWhite
txt.ForeColor = vbBlack
End Sub
Public Sub TextKeyD(KeyCode As Integer)
If KeyCode = 37 Then
SendKeys "+{tab}"
KeyCode = 0
End If
End Sub
Public Sub Pcase(KeyAscii As Integer, Ptxt As TextBox)
'This Function Will write Text in Proper case
If Ptxt.SelStart = 0 Then
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
Else
If Mid$(Ptxt, Ptxt.SelStart, 1) = Space$(1) Then
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End If
End If
End Sub
    

