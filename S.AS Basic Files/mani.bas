Attribute VB_Name = "Module1"
Public Sub main()
On Error Resume Next
If Form1.Msn.Services.PrimaryService.Status = MSS_LOGGED_ON Then
Call En
End If
If Form1.Msn.Services.PrimaryService.Status = MSS_NOT_LOGGED_ON Then
Call DN
MsgBox "Please Sign In To Use The Micro Nick Machine", vbExclamation
End If
End Sub
Public Sub En()
On Error Resume Next
Form1.Frame1.Enabled = True
Form1.Frame2.Enabled = True
Form1.Frame3.Enabled = True
Form1.Frame4.Enabled = True
Form1.Frame5.Enabled = True
Form1.Frame6.Enabled = True
Form1.Frame7.Enabled = True
Form1.Frame8.Enabled = True
Form1.Frame9.Enabled = True
Form1.Frame10.Enabled = True
End Sub
Public Sub DN()
On Error Resume Next
Form1.Frame1.Enabled = False
Form1.Frame2.Enabled = False
Form1.Frame3.Enabled = False
Form1.Frame4.Enabled = False
Form1.Frame5.Enabled = False
Form1.Frame6.Enabled = False
Form1.Frame7.Enabled = False
Form1.Frame8.Enabled = False
Form1.Frame9.Enabled = False
Form1.Frame10.Enabled = False
End Sub
