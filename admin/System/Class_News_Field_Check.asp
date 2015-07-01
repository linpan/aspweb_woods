<!--#include file="../../Include/Class_Conn.asp"-->
<!--#include file="../../Include/Class_Main.asp"-->
<% Super=1 '只允许超级管理员访问该页 %>
<!--#include file="../p8_Check.asp"-->
<%
	Response.Buffer=True 
	Response.ExpiresAbsolute=Now()-1 
	Response.Expires=0 
	Response.CacheControl ="no-cache" 
	Response.AddHeader "Pragma","No-Cache"

	Dim Variable,rs
	Variable = Lcase(Request.Querystring("Variable"))
	ClassNum = Lcase(Request.Querystring("ClassNum"))
	Set rs = Server.CreateObject ("ADODB.Recordset")
	rs.Open "Select id From p8_Field Where ClassNum = '"& ClassNum &"' And Variable='"& Variable &"'",Conn,1,1

	If Not rs.Eof Then
		Response.Write "1"
	Else
		Response.Write "0"
	End If
	
	rs.Close
	Set rs=Nothing
	CloseConn
%>

