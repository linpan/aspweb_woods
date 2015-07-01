<%
	Dim Admin_UserName,Admin_PassWord,Admin_rs,Super
	If Request.Cookies("Admin")("s_User")<>"" And Request.Cookies("Admin")("s_Pass")<>"" Then
		Admin_UserName = Replace_Text(Request.Cookies("Admin")("s_User"))
		Admin_PassWord = Replace_Text(Request.Cookies("Admin")("s_Pass"))
		
		Set Admin_rs = Server.CreateObject ("ADODB.Recordset")
		Admin_rs.Open "Select s_Level From p8_Super Where s_User='"& Admin_UserName &"' And s_Pass='"& Admin_PassWord &"'",Conn,1,1
	
		If Admin_rs.Eof Then
			If ExitPath = 1 Then
				Response.Write "<script>top.location.href='Index.asp';</script>"
			Else
				Response.Write "<script>top.location.href='../Index.asp';</script>"
			End If
		Else
			If Super = 1 Then
				If Admin_rs("s_Level") <> 1 Then
					Response.Write "您无权使用该功能！" 
					Response.End()
				End If
			End If
		End If
		
		Admin_rs.Close
		Set Admin_rs = Nothing
	Else
		If ExitPath = 1 Then
			Response.Write "<script>top.location.href='Index.asp';</script>"
		Else
			Response.Write "<script>top.location.href='../Index.asp';</script>"
		End If
	End If
%>