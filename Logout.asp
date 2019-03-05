


<%
	
	Session.Contents.Remove("username")
	
	Session.Abandon()
	
	Response.Redirect"Login.asp"

%>