'Sample script to try the NLSProcReq DLL

Set NLS = CreateObject("NLSProcReq.ExposedFunctions")
NLS.ConnectionName = "CN"
NLS.UserName = "UN"
NLS.Password = "PW"
   
MsgBox NLS.GetNLSProcReqVersion
MsgBox NLS.GetNLSProcReqLocation
    
NLS.InitializedConnection

MsgBox "ErrorMessage : " + NLS.ErrorMessage
