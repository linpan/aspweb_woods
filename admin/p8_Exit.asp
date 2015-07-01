<%
	Response.Cookies("Admin")("s_User") = ""
	Response.Cookies("Admin")("s_Pass") = ""
	Response.Redirect "Index.asp"
%>