
If (nlsapp.GetCurrentUserNo=0 or nlsapp.GetCurrentUserNo=1) Then
     ' proceed
Else
     MsgBox "You are not allowed to proceed."
     nlsapp.break
End If
