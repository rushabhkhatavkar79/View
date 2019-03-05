<html>
<head></head>
<body>
<h2 align=center>User Information</h2>
<form action="Logout.asp">
<div align="center">
<table border="3" cellpadding="4" cellspacing="5">
<%

Dim SQLQuery

Dim rs

Dim UserName

Dim Password


	UserName = request.form("txtusername")
	Password = request.form("txtpassword")
	If UserName <> "" and Password <> "" Then

		set Con=CreateObject("ADODB.Connection")
		set rs =CreateObject("ADODB.Recordset")
		Con.ConnectionString="Provider=SQLOLEDB;Data-Source=IDTP314;Database=Demo;User ID=sa;password=Synerz!p"
		Con.open

		SQLQuery = "select * from UserInfo where uname = '"&UserName&"' AND pwd = '"&Password&"'"
	
		set rs  = Con.execute(SQLQuery)

		If rs.BOF or rs.EOF Then
	
			Response.Redirect "login.asp"
		Else
	
			Response.Cookies("UserName")=UserName
			Response.Cookies("Password")=Password
		
			SQLQuery="create view information as select * from UserInfo"
			SQLQuery="select * from information"
			set rs  = Con.execute(SQLQuery)
			
						
						Response.Write("<tr>")
						
						Response.Write ("<th>") 
						Response.Write"Name"
						Response.Write("</th>")
						
						Response.Write ("<th>") 
						Response.Write"Username"
						Response.Write("</th>")
						
						Response.Write ("<th>") 
						Response.Write"Password"
						Response.Write("</th>")
						
						Response.Write ("<th>") 
						Response.Write"Mobile"
						Response.Write("</th>")
						
						Response.Write ("<th>") 
						Response.Write"Address"
						Response.Write("</th>")
						
						'Response.Write("</tr>")
						
						
						
						Response.Write("</tr>")
						
						do until rs.EOF
			
				
						Set objFields = rs.Fields
						Response.Write("<tr>")
						For intLoop = 0 To (objFields.Count - 1)
						Response.Write("<td>")
						Response.Write(objFields.Item(intLoop).Value)
						Response.Write("</td>")
						
						
						
						
						
						Next  
							Response.Write("</tr>")
					
				
							rs.MoveNext
		
						loop
						
				Response.Write("</table>")
		End If
		SQLQuery="drop view information"
		Con.Close
		
		If rs.State = adStateOpen then 
	
			'rs.Close
	
		End If
	
		set rs = nothing
		set Con = nothing

	End If


%>

<table border="0" cellpadding="3" cellspacing="3">

<tr>
<td></td>
	<td><input type="submit" value="Logout" style="font-size:15pt">
</td>
</tr>
</table>
</div>
</body>
</html>