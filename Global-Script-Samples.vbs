'******** Encryption************
Encrypt_Key = "GGIT***@@jkJKL"

Function Ecrypt(str,key)
           Dim lenKey, KeyPos, LenStr, x, Newstr
 
           Newstr = ""
           lenKey = Len(key)
           KeyPos = 1
           LenStr = Len(Str)
           str = StrReverse(str)
           For x = 1 To LenStr
                Newstr = Newstr & chr(asc(Mid(str,x,1)) + Asc(Mid(key,KeyPos,1)))
                KeyPos = keypos+1
                If KeyPos > lenKey Then KeyPos = 1
            Next
           encrypt = Newstr
End Function
Function Decrypt(str,key)
           Dim lenKey, KeyPos, LenStr, x, Newstr
 
           Newstr = ""
           lenKey = Len(key)
           KeyPos = 1
           LenStr = Len(Str)
           str=StrReverse(str)
           For x = LenStr To 1 Step -1
                Newstr = Newstr & chr(asc(Mid(str,x,1)) - Asc(Mid(key,KeyPos,1)))
                KeyPos = KeyPos+1
                If KeyPos > lenKey Then KeyPos = 1
           Next
           Newstr=StrReverse(Newstr)
           Decrypt = Newstr
End Function

'******** Encryption************


'**********Connection using Encryption*******

Function globalEncrptCreateNLSObject()
         str = "Ê¼ÏÂÂ"
         Set  NLSlocal = CreateObject("NLSProcReq.ExposedFunctions")
         NLSlocal.ConnectionName = "xxxxxx" 
         NLSlocal.Password =   Decrypt(str,Encrypt_Key)
         NLSlocal.Username =  Decrypt(str,Encrypt_Key)
         Set globalEncrptCreateNLSObject = NLSlocal
End Function
'**********Connection using Encryption*******


'*******New Task XML****
Function GLOBAL_GenerateTaskXML(iTaskTemplate, strStatus, strSubject, strPriority)
	GLOBAL_GenerateTaskXML = "<TASK" & _
                                " TaskTemplateNo = """ &  iTaskTemplate & """" & _
                                " UpdateFlag=""0""" & _
                                " StatusCodeName = """ &  strStatus & """" & _
		" PriorityCodeName = """ &  strPriority & """" & _
		" Subject  = """ &  strSubject & """" & _
		" />" & vbCrLf
End Function
'*******New Task XML****

Function C_Cdbl(val)
	If IsNumeric(val) Then
		C_Cdbl = Cdbl(val)
	Else
		C_Cdbl = 0.0
	End If
End Function

Function C_Cint(val)
	If IsNumeric(val) Then
		C_Cint = Clng(val)
	Else
		C_Cint = 0
	End If
End Function

Function C_CDate(val, default)
	If IsDate(val) Then
		C_CDate = CDate(val)
	Else
		C_CDate = default
	End If
End Function

Function C_CRound(val, iPrecision)
	C_CRound = 0.0
	If IsNumeric(val) AND IsNumeric(iPrecision) Then
		C_CRound = Round(Cdbl(val), Cint(iPrecision))
	End If
End Function

Sub Append(ByRef str, strDelimit, strApp)
	If str <> "" Then
		str = str & strDelimit
	End If
	str = str & strApp
End Sub

Function FormatXMLDate(strDate)
	dtDate = Date
	If IsDate(strDate) Then
		dtDate = CDate(strDate)
	End If
	FormatXMLDate = Right("0" & Month(dtDate), 2) & "/" & Right("0" & Day(dtDate), 2) & "/" & Year(dtDate)
End Function

Function FormatDateField(strDate)
	dtDate = Date
	If IsDate(strDate) Then
		dtDate = CDate(strDate)
	End If
	FormatDateField = Year(dtDate) & "/" & Right("0" & Month(dtDate), 2) & "/" & Right("0" & Day(dtDate), 2)
End Function

Function XMLEncode(strText)
	XMLEncode = Replace(Replace(Replace(Replace(strText, "&", "&amp;"), ">", "&gt;"), "<", "&lt;"), """", "&quot;")
End Function

Function GLOBAL_CreateNLSObject()
	If VarType(NLSapp) = 0 Then 'Web Service
		Err.Clear
		On Error Resume Next
		Set NLS = GLOBAL_CreateNLSProcReq2()
		Set GLOBAL_CreateNLSObject = NLS

	Else 'GUI
		Set GLOBAL_CreateNLSObject = NLSapp
	End If
End Function

Function GLOBAL_CreateNLSProcReq(strServer, strDatabase, strUsername)
	Set l_NLS = CreateObject("NLSProcReq.ExposedFunctions")
	l_NLS.SetWindowsAuthenticationCredential strServer, strDatabase, "ORACLE", strUsername

	Set GLOBAL_CreateNLSProcReq = l_NLS
End Function

Function GLOBAL_CreateNLSProcReq2()
	Set GLOBAL_CreateNLSProcReq2 = GLOBAL_CreateNLSProcReq(CGlobalNLSServer, CGlobalNLSDatabase, CGlobalNLSSignonName)
End Function

Function GLOBAL_ImportXML(ByRef l_NLS, strXML, ByRef strOutput)
	GLOBAL_ImportXML = GLOBAL_ImportXML2(l_NLS, strXML, strOutput, True, False)
End Function

Function GLOBAL_ImportXML2(ByRef l_NLS, strXML, ByRef strOutput, bSaveRequired, bForce)
	strOutput = ""
	GLOBAL_ImportXML2 = False
	If Not IsGUI() OR bForce Then
		If VarType(l_NLS) = 0 Then
			l_NLS = GLOBAL_CreateNLSProcReq2()
		End If

		l_NLS.ImportString = strXML
		l_NLS.ErrorMessage = ""
		GLOBAL_ImportXML2 = l_NLS.ImportXML
		strOutput = l_NLS.ErrorMessage
	Else
		If bSaveRequired AND NLSapp.IsCurrentViewDirty() Then
			NLSapp.SaveScreen()
		End If
		NLSapp.ClearErrorString()
		GLOBAL_ImportXML2 = NLSapp.ImportXMLRecord(strXML)
		strOutput = NLSapp.GetErrorString()
	End If
End Function

Function GLOBAL_WrapXML(strXML)
	GLOBAL_WrapXML = GLOBAL_WrapXML2(strXML, 1)
End Function

Function GLOBAL_WrapXML2(strXML, iCommitBlock)
	If iCommitBlock <> 1 Then
		iCommitBlock = 0
	End If
	GLOBAL_WrapXML2 = "<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>" & vbCrLf & _
		"<NLS EnforceTagExistence=""1"" CommitBlock=""" & iCommitBlock & """ >" & vbCrLf & _
			strXML & vbCrLf & _
		"</NLS>"
End Function

Function GLOBAL_WrapTransactionXML(strXML)
	GLOBAL_WrapTransactionXML = ""& _
		"<TRANSACTIONS>" & vbCrLf & _
			strXML & _
		"</TRANSACTIONS>" & vbCrLf
End Function

Function GLOBAL_GenerateTransactionCodeXML(iTransCode, dtEffDate, strLoanNum, dAmount)
	GLOBAL_GenerateTransactionCodeXML = "" & _
		"<TRANSACTIONCODE" & _
		" TransactionCode=""" & iTransCode & """" & _
		" EffectiveDate=""" & FormatXMLDate(dtEffDate) & """" & _
		" LoanNumber=""" & strLoanNum & """" & _
		" Amount=""" & C_CRound(dAmount, 2) & """" & _
		" />" & vbCrLf
End Function

Function GLOBAL_GenerateLoanCommentXML(strComment, strCommentDesc, strCategory)
	strXML = "<LOANCOMMENTS" & _
			" Comment=""" & XMLEncode(strComment) & """" & _
			" CommentDescription=""" & XMLEncode(strCommentDesc) & """" & _
			" Category=""" & XMLEncode(strCategory) & """" & _
			" />" & vbCrLf
	GLOBAL_GenerateLoanCommentXML = strXML
End Function

Function GLOBAL_AutomatedPaymentsActive(l_NLS, iAcctRefno)
	bFlag = False
	strSQL = "SELECT 1" & _
			" FROM loanacct_ach ach" & _
			" WHERE ach.status = 0" & _
			" AND ach.acctrefno = " & iAcctRefno
	If C_Cint(NLSapp.SQLSelectStatement(strSQL)) = 1 Then
		bFlag = True
	End If
	GLOBAL_AutomatedPaymentsActive = bFlag
End Function

Function GLOBAL_SendEmail(l_NLS, strEmailTo, strEmailFrom, strSubject, strMessage)
	GLOBAL_SendEmail = True
	l_NLS.ErrorMessage = ""
	If Not l_NLS.SendEmailMessage(cGlobalEmailServer, strEmailTo, strEmailFrom, strSubject, strMessage, "") Then
		GLOBAL_SendEmail = False
		GLOBAL_LogError l_NLS, _
			"Error sending email: " & l_NLS.ErrorMessage & vbCrlf & _
			vbTab & "To: " & strEmailTo & vbCrlf & _
			vbTab & "From: " & strEmailFrom & vbCrlf & _
			vbTab & "Subject: " & strSubject & vbCrlf & _
			vbTab & "Message: " & strMessage
	End If
End Function

Function IsGUI()
	If VarType(NLSapp) = 0 Then
		IsGUI = False
	Else
		IsGUI = True
	End If
End Function

Sub GLOBAL_LogError(l_NLS, strMessage)
	GLOBAL_Log l_NLS, strMessage, "Scripting", 1
End Sub

Sub GLOBAL_Log(l_NLS, strMessage, strLogType, iErrorStatus)
	strMsg = XMLEncode(strMessage)
	strXML = "<NLS>" & vbCrLf & _
				"<NLSLOG" & _
				" LogTypeDescription=""" & strLogType & """" & _
				" ErrorStatus=""" & iErrorStatus & """" & _
				" Log=""" & strMsg & """" & _
				" />" & vbCrLf & _
			"</NLS>" & vbCrLf
	GLOBAL_ImportXML l_NLS, strXML, ""
End Sub

Function Global_GenerateLoanStatusXML(strStatus, strOp, strDate)
	strEffDate = FormatXMLDate(strDate)
	strXML = "<LOANSTATUSES" & _
			" LoanStatusCode=""" & strStatus & """" & _
			" Operation=""" & strOp & """" & _
			" EffectiveDate=""" & strEffDate & """" & _
			" />" & vbCrLf
	Global_GenerateLoanStatusXML = strXML
End Function
