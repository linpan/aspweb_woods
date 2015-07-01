<!--#include file="../../Include/Class_Conn.asp"-->
<!--#include file="../../Include/Class_MD5.asp"-->
<!--#include file="../p8_Check.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title></title>
<link href="../css/Public.css" rel="stylesheet" type="text/css" />
<meta http-equiv="pragma" content="no-cache" />
<meta http-equiv="Cache-Control" content="no-cache, must-revalidate" />
<meta http-equiv="expires" content="Wed, 26 Feb 1997 08:21:57 GMT" />
<meta http-equiv="expires" content="0" />
<style type="text/css">
<!--
	.Btn {padding-left:15px; background:url(../images/ico_4.gif) 6px 8px no-repeat;}
	.Btn_over { padding-left:15px; background:url(../images/ico_4.gif) #eaf3fd 6px 8px no-repeat;}
-->
</style>
</head>

<body>
<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="120" valign="top" bgcolor="#EAF3FD"><table width="100%" border="0" cellspacing="1" cellpadding="0">
      <tr>
        <td height="25" align="center" bgcolor="#5B92DB"><font color="#FFFFFF"><strong>文章管理</strong></font></td>
      </tr>
      <%
	Set rs = Server.CreateObject("Adodb.Recordset")
	rs.Open "Select id,ClassName From p8_Class Where ClassType = '文章' And ClassLevel=1 Order By id Desc",Conn,1,1
	
	Do While Not rs.Eof
		Response.Write "<tr><td class=""Btn"" height=""25""> <a href=""News_List.asp?ClassID="& rs("id") &""" target=""smain"" title="""& rs("ClassName") &""">"& Left(rs("ClassName"),4) &"</a> |  <a href=""News_Add.asp?ClassID="& rs("id") &""" target=""smain"">添加</a></td></tr>"
		rs.MoveNext
	Loop
	
	rs.Close
	Set rs = Nothing
%>
      <tr>
        <td height="25" align="center" bgcolor="#5B92DB"><font color="#FFFFFF"><strong>图片管理</strong></font></td>
      </tr>
      <%
	Set rs = Server.CreateObject("Adodb.Recordset")
	rs.Open "Select id,ClassName From p8_Class Where ClassType = '图片' And ClassLevel=1 Order By id Desc",Conn,1,1
	
	Do While Not rs.Eof
		Response.Write "<tr><td class=""Btn"" height=""25""> <a href=""Pic_List.asp?ClassID="& rs("id") &""" target=""smain"" title="""& rs("ClassName") &""">"& Left(rs("ClassName"),4) &"</a> |  <a href=""Pic_Add.asp?ClassID="& rs("id") &""" target=""smain"">添加</a></td></tr>"
		rs.MoveNext
	Loop
	
	rs.Close
	Set rs = Nothing
%>
      <tr>
        <td height="25" align="center" bgcolor="#5B92DB"><font color="#FFFFFF"><strong>单页管理</strong></font></td>
      </tr>
	<%
		Set rs = Server.CreateObject("Adodb.Recordset")
		rs.Open "Select id,ClassName From p8_Class Where ClassType = '单页' And ClassLevel=1 Order By id Desc",Conn,1,1
		
		Do While Not rs.Eof
			Response.Write "<tr><td class=""Btn"" height=""25""> <a href=""Page_Edit.asp?ClassID="& rs("id") &"&ClassName="& rs("ClassName") &""" target=""smain"" title="""& rs("ClassName") &""">"& Left(rs("ClassName"),7) &"</a></td></tr>"
			rs.MoveNext
		Loop
		
		rs.Close
		Set rs = Nothing
	%>
      <tr>
        <td height="25" align="center" bgcolor="#5B92DB"><font color="#FFFFFF"><strong>表单管理</strong></font></td>
      </tr>
	<%
		Set rs = Server.CreateObject("Adodb.Recordset")
		rs.Open "Select id,ClassName From p8_Class Where ClassType = '表单' And ClassLevel=1 Order By id Desc",Conn,1,1
		
		Do While Not rs.Eof
			Response.Write "<tr><td class=""Btn"" height=""25""> <a href=""Form_List.asp?ClassID="& rs("id") &""" target=""smain"" title="""& rs("ClassName") &""">"& Left(rs("ClassName"),9) &"</a></td></tr>"
			rs.MoveNext
		Loop
		
		rs.Close
		Set rs = Nothing
	%>
      <tr>
        <td height="25" align="center" bgcolor="#5B92DB"><font color="#FFFFFF"><strong>在线客服</strong></font></td>
      </tr>
	  <tr><td class="Btn" height="25"> <a href="Service.asp" target="smain">设置在线客服</a></td></tr>
    </table></td>
    <td valign="top"><iframe width="100%" frameborder="0" id="smain" name="smain" src="News_List.asp" style="height:100%;" scrolling="yes"></iframe></td>
  </tr>
</table>
</body>
</html>

