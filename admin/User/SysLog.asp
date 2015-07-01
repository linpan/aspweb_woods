<!--#include file="../../Include/Class_Conn.asp"-->
<!--#include file="../../Include/Class_Main.asp"-->
<% Super=1 '只允许超级管理员访问该页 %>
<!--#include file="../p8_Check.asp"-->
<%
	Dim rs,Page,UserName,UserName_Sql
	UserName = Request("UserName")
	
	If UserName <>"" Then
		UserName_Sql = " And UserName Like '%"& UserName &"%' "
	End If
	
	Set rs = Server.CreateObject("ADODB.RecordSet")
	rs.open "Select * From p8_Log Where 1=1 "& UserName_Sql & " Order By id Desc",Conn,1,1
	
	rs.PageSize = 20
	If Request("Page") <> "" Then
		Page = Cint(Request("Page"))
	Else
		Page = 1
	End If
	If Not rs.Eof And Not rs.Bof Then
		rs.AbsolutePage = Page
	End If
	Sum = rs.PageSize
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>后台登录日志</title>
<script type="text/javascript">top.window.aTitle.innerText='后台登录日志'</script>
<meta http-equiv="pragma" content="no-cache" />
<meta http-equiv="Cache-Control" content="no-cache, must-revalidate" />
<meta http-equiv="expires" content="Wed, 26 Feb 1997 08:21:57 GMT" />
<meta http-equiv="expires" content="0" />
<link href="../css/Public.css?admin" rel="stylesheet" type="text/css" /> 
</head>
<script type="text/javascript" src="../Include/TipBox.js"></script>
<script type="text/javascript" src="../Include/calendar.js"></script>
<body>
<%
	Dim Tip
	If Tip = "" Then Tip = Request.QueryString("Tip")
	If Tip <> "" Then
		Response.Write "<script type=""text/javascript"">window.onload=function(){new x.creat(1, 41, 5, 10, '"& Tip &"');}</script>"
	End If
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="64%" height="30" bgcolor="#eaf3fd" style="border-bottom:1px solid #b5cef0;">
	<form name="form1" method="post" action="SysLog.asp">
	
      <table style="margin-left:10px;" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="140">用户名：<input style="width:80px;" name="UserName" type="text" <%If UserName="" Then Response.Write "class=""ipt2""" :Else Response.Write "class=""ipt"""%> value="<%=UserName%>" maxlength="50"></td>
          <td><input name="Submit" type="submit" class="btn1" value="搜索">
            &nbsp;&nbsp;<input name="Submit" type="button" class="btn1" onClick="window.location.href='SysLog.asp'" value="全部"></td>
          </tr>
      </table>
    </form>    </td>
    <td width="36%" align="right" bgcolor="#eaf3fd" style="border-bottom:1px solid #b5cef0;"><font color="#FF0000">提示：系统只保留30天内日志</font>&nbsp;&nbsp;</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="1" cellpadding="0">
  <tr bgcolor="#E4EDF9">
    <td width="17%" height="25" align="center">用户名</td>
    <td width="18%" align="center">用户权限</td>
	<td width="21%" align="center">登录时间</td>
	<td width="26%" align="center">登录IP</td>
    <td width="18%" align="center">事件</td>
  </tr>
<%
If rs.RecordCount = 0 Then
	Response.Write "<tr bgcolor=""#F8FBFE""><td height=""400"" colspan=""8"" align=""center"">没有找到相关信息！</td></tr>"
Else
	Do While Not rs.Eof And Sum>0 
%>
	<tr onMouseOver="this.style.backgroundColor='#e7ffa6'" onMouseOut="this.style.backgroundColor='#F8FBFE'" bgcolor="#F8FBFE">
	<td height="25" align="center"><%=rs("UserName")%></td>
	<td align="center"><%
	If rs("UserLevel") = 1 Then
		Response.Write "超级管理员"
	Else
		Response.Write "录入员"
	End If
	%></td>
	<td align="center"><%=rs("LoginIP")%></td>
	<td align="center"><%=rs("LoginDate")%></td>
	<td align="center"><%
	If Instr(rs("LoginState"),"失败") Then
		Response.Write "<font color=""red"">"& rs("LoginState") &"</font>"
	Else
		Response.Write rs("LoginState")
	End If
	%></td>
	</tr>
<%
	rs.MoveNext     
	Sum = Sum - 1     
	Loop 
End If
%>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <form name="Page" method="Post" action="SysLog.asp">
      <td height="50" align="center" valign="middle" bordercolor="#FFFFFF">共<font color="#FF2D00"><%=rs.RecordCount%></font>条&nbsp;&nbsp;<font color="#FF2D00"><%=Page%></font>/<font color="#FF2D00"><%=rs.pagecount%></font>&nbsp;&nbsp;
	  <a href="?Page=1&UserName=<%=UserName%>" class="Text_1">首页</a>
          <%If Page>1 Then%>
          <a href="?Page=<%=Page-1%>&UserName=<%=UserName%>" class="Text_1">上一页</a>
          <%else%>
        上一页
        <%End If%>
        <%If Page < rs.pagecount Then %>
        <a href="?Page=<%=Page+1%>&UserName=<%=UserName%>" class="Text_1">下一页</a>
        <%else%>
        下一页
        <%End If%>
        <a href="?Page=<%=rs.pagecount%>&UserName=<%=UserName%>" class="Text_1">尾页</a>
        <input name="Page" type="text" class="ipt2" id="Page" value="<%=Page%>" size="3">
        <input name="Submit2" type="submit" class="ipt2" value="GO">
        <input name="UserName" type="hidden" value="<%=UserName%>">
	</td>
    </form>
  </tr>
</table>
</body>
</html>
<%
	rs.close
	Set rs=Nothing
	CloseConn
%>