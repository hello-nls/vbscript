'Sample script to try the NLSProcReq DLL 
'using an external vbScript file saved at desktop 

Set NLS = CreateObject("NLSProcReq.ExposedFunctions")
NLS.ConnectionName = "CN"
NLS.UserName = "UN"
NLS.Password = "PW"
   
MsgBox NLS.GetNLSProcReqVersion
MsgBox NLS.GetNLSProcReqLocation
    
NLS.InitializedConnection

MsgBox "ErrorMessage : " + NLS.ErrorMessage
