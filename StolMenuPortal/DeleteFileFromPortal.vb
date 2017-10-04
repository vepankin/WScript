strURL = "http://portal/SiteCollectionDocuments/"

Set HTTP = CreateObject("MSXML2.XMLHTTP")
	
	HTTP.open "DELETE", strURL & "Test.txt", False 
	
	HTTP.send 

Set HTTP = Nothing
