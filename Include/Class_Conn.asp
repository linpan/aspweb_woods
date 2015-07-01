<!--#include file="Class_Safe.asp" -->
<%
Response.Buffer = True
Dim Conn

Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open "Provider=MicroSoft.Jet.Oledb.4.0; Data Source=" & Server.Mappath("/Psd8_Data/#psd8_com-3i58g5.mdb")

Function CloseConn
	Conn.close
	Set Conn = Nothing
End Function

Function CloseRs
	Rs.Close
	Set Rs = Nothing
End Function

If Err Then
	err.Clear
	Set Conn = Nothing
	Response.Write "数据库连接出错，请检查连接字串。"
	Response.End
End If
%>