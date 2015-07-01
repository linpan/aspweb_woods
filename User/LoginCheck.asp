<%
	If Session("UserName")  = "" Or Session("UserState") = "" Then
		Response.Redirect "Login.asp"
		Response.End()
	End If
%>