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

	Dim id,rs
	id = Request.Querystring("id")
	
	Conn.Execute("Delete From p8_Field Where id = " & id )
	Response.Write "1"
	Response.End()
	
	rs.Close
	Set rs=Nothing
	CloseConn
%>

