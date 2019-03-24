/*
 * Source: https://community.nortridge.com/t/getting-the-last-4-of-social/2298/25  
 *                 From stanleyM
 * /

'Generic decryption function for any NLS encrypted data from stanleyM
Function DecryptNLS(x)

  dekey = "[YOUR ENCRYPTION KEY HERE]"

  'Format validation
   xvalidate = left(x,9) & mid(x,32,3) & right(x,2)

   If xvalidate = "[000001][==]==" Then
      Encrpt = right(x,24)
      EncrptIV = mid(x,10,24)
      DecryptNLS = NLSApp.NLSAESDecrypt(Encrpt, dekey, EncrptIV)
   Else
      DecryptNLS = "NA"
   End if

End function





' -------------------------------------






/*
 * sample test script from Jojo
 */

testVal = DecryptNLS("SAMPLE TEST ENCRYPTED DATA")
last4ss = Mid(testval, len(testval)-3 )
msgbox "ss: " & testval & "  - - >  last4ss: " & last4ss


       Function DecryptNLS(x)

          dekey = "[YOUR ENCRYPTION KEY HERE]"               

         'Format validation
          xvalidate = left(x,9) & mid(x,32,3) & right(x,2)
          
         'msgbox xvalidate

          If xvalidate = "[000001][==]==" Then
          
              Encrpt = right(x,24)
              'msgbox Encrpt

              EncrptIV = mid(x,10,24)
              'msgbox EncrptIV

              DecryptNLS = NLSApp.NLSAESDecrypt(Encrpt, dekey, EncrptIV)

          Else          
              DecryptNLS = "NA"              
          End if

      End function


' -------------------------------------


