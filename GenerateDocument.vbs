/*
 * Source: https://community.nortridge.com/t/got-any-dealbreaker-showstopper-info-issues-concerns-re-v5-13-0-v5-13-1/2260/24 
 */


Set NLS = CreateObject("NLSProcReq.ExposedFunctions")
NLS.ConnectionName = NLSapp.GetConnectionName()
UserNoX = CInt(NLSapp.GetCurrentUserNo())
NLS.UserName =  NLSapp.SQLSelectStatement("SELECT signonid FROM nlsusers where userno = " & UserNoX)


NLS.Password = NLSapp.GetSignedOnUserPasswordEncrypted()
'msgbox NLS.Password

NLS.Password = "HARDCODEDPASSWORD"

IF NLS.InitializedConnection() then

   msgbox "Initialized"

   loanrefno = nlsapp.getfield("LOAN_REFNO")

   SET DT = NLS.GetDocumentTemplates(29)
   DT.GenerateDocument loanrefno , 1, "", 0, ""

end if

MsgBox "ErrorMessage : " + NLS.ErrorMessage
