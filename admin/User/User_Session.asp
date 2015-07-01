<!--#include file="../../Include/Class_Conn.asp"-->
<!--#include file="../../Include/Class_Main.asp"-->
<!--#include file="../p8_Check.asp"-->
<%
	Dim UserName,rs
	UserName  = Request("UserName")
	
	Set rs = Server.CreateObject ("ADODB.Recordset")
	rs.Open "Select UserName,UserState From p8_User Where UserName='"& UserName &"'",conn,1,1
	
	If Not rs.Eof Then
		Session("UserName")  = rs("UserName")
		Session("UserState") = rs("UserState")
		
		Response.Redirect "../../User/Main.asp"
		Response.End()
	Else
		Response.Write "<script>alert('用户不存在');window.close();</script>"
	End If
%>