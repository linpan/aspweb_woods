<!--#include file="../Include/Class_Conn.asp"-->
<!--#include file="../Include/Class_Main.asp"-->
<%
	Session("UserName")  = ""
	Session("UserState") = ""
	Response.Redirect SiteDir
%>