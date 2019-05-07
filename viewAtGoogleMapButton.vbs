   cifno = nlsapp.GetField("CIF_REFNO")
   URL = nlsapp.SQLSelectStatement("select REPLACE(street_address1+','+city+','+zip,' ','+') from cif WHERE cifno="+cifno)
   
   dim ie
   set ie = createobject("InternetExplorer.Application")
   ie.Navigate "http://maps.google.com/maps?q="+URL
   ie.Visible = true

