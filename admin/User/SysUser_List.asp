<!--#include file="../../Include/Class_Conn.asp"-->
<!--#include file="../../Include/Class_Main.asp"-->
<% Super=1 '只允许超级管理员访问该页 %>
<!--#include file="../p8_Check.asp"-->
<%
'删除============================================================================================
If Request.QueryString("ac")="del" Then
	id      = Request("id")
	Page    = Request("Page")
	s_Level = Request("s_Level")
	s_User  = Request("s_User")
	s_Name  = Request("s_Name")
	
	Set rs = Server.Createobject("Adodb.RecordSet")
	rs.open "Select id From p8_Super Where id= " & id ,Conn,1,3
	
	rs.Delete
	rs.Close
	Set rs = Nothing
	Response.Redirect "SysUser_List.asp?Tip=删除成功！&Page="& Page &"&s_Level="& s_Level &"&s_User="& s_User &"&s_Name="& s_Name
	Response.End()
End If
'/删除============================================================================================

	Dim rs,Page,s_Level,s_User,s_Name,s_Level_Sql,s_User_Sql,s_Name_Sql
	s_Level = Request("s_Level")
	s_User  = Request("s_User")
	s_Name  = Request("s_Name")
	
	If s_Level <>"" Then
		s_Level_Sql = " And s_Level = "& s_Level
	End If

	If s_User <>"" Then
		s_User_Sql = " And s_User Like '%"& s_User &"%' "
	End If
	
	If s_Name <>"" Then
		s_Name_Sql = " And s_Name Like '%"& s_Name &"%' "
	End If
	
	Set rs = Server.CreateObject("ADODB.RecordSet")
	rs.open "Select * From p8_Super Where 1=1 "& s_Level_Sql & s_User_Sql & s_Name_Sql &" Order By id Desc",Conn,1,1
	
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
<title>网站管理员</title>
<script type="text/javascript">top.window.aTitle.innerText='网站管理员'</script>
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
	Tip = Request.QueryString("Tip")
	If Tip <> "" Then
		Response.Write "<script type=""text/javascript"">window.onload=function(){new x.creat(1, 41, 5, 10, '"& Tip &"');}</script>"
	End If
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="30" bgcolor="#eaf3fd" style="border-bottom:1px solid #b5cef0;">
	<form name="form1" method="post" action="SysUser_List.asp">
	
      <table style="margin-left:10px;" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="110">
		    <select name="s_Level" id="s_Level" class="ipt5">
		      <option value="">权限</option>
			  <option value="1" <%If s_Level = "1" Then Response.Write " selected=""selected"""%>>超级管理员</option>
		      <option value="2" <%If s_Level = "2" Then Response.Write " selected=""selected"""%>>录入员</option>
		    </select>			</td>
          <td width="160">用户名：<input style="width:100px;" name="s_User" type="text" <%If s_User="" Then Response.Write "class=""ipt2""" :Else Response.Write "class=""ipt"""%> value="<%=s_User%>" maxlength="50"></td>
          <td width="140">姓名：
            <input style="width:80px;" name="s_Name" type="text" <%If s_Name="" Then Response.Write "class=""ipt2""" :Else Response.Write "class=""ipt"""%> value="<%=s_Name%>" maxlength="50"></td>
          <td><input name="Submit" type="submit" class="btn1" value="搜索">
            &nbsp;&nbsp;<input name="Submit" type="button" class="btn1" onClick="window.location.href='SysUser_List.asp'" value="全部"></td>
          </tr>
      </table>
    </form>    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="1" cellpadding="0">
  <tr bgcolor="#E4EDF9">
    <td width="16%" height="25" align="center">权限</td>
    <td width="13%" align="center">用户名</td>
	<td width="11%" align="center">姓名</td>
    <td width="16%" align="center">最后登录IP</td>
    <td width="21%" align="center">最后登录时间</td>
    <td width="9%" align="center">登录次数</td>
    <td width="14%" align="center">操作</td>
  </tr>
<%
If rs.RecordCount = 0 Then
	Response.Write "<tr bgcolor=""#F8FBFE""><td height=""400"" colspan=""8"" align=""center"">没有找到相关信息！</td></tr>"
Else
	Do While Not rs.Eof And Sum>0 
%>
	<tr onMouseOver="this.style.backgroundColor='#e7ffa6'" onMouseOut="this.style.backgroundColor='#F8FBFE'" bgcolor="#F8FBFE">
	<td height="25" align="center">
	<%
		If rs("s_Level") = "1" Then
			Response.Write "超级管理员"
		End If
		If rs("s_Level") = "2" Then
			Response.Write "录入员"
		End If
	%>
	</td>
	<td align="center">
	  <a href="SysUser_Edit.asp?id=<%=rs("id")%>&Page=<%=Page%>&s_Level=<%=s_Level%>&s_User=<%=s_User%>&s_Name=<%=s_Name%>"><%=rs("s_User")%></a>
	</td>
	<td align="center"><%=rs("s_Name")%></td>
	<td align="center"><%=rs("s_IP")%></td>
	<td align="center"><%=rs("s_Date")%></td>
	<td align="center"><%=rs("s_Count")%></td>
	<td align="center">
	 <a href="SysUser_Edit.asp?id=<%=rs("id")%>&Page=<%=Page%>&s_Level=<%=s_Level%>&s_User=<%=s_User%>&s_Name=<%=s_Name%>">修改</a> <a href="javascript:if(confirm('删除后不可恢复，是否继续？'))window.location.href='?ac=del&id=<%=rs("id")%>&Page=<%=Page%>&s_Level=<%=s_Level%>&s_User=<%=s_User%>&s_Name=<%=s_Name%>';">删除</a></td>
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
    <form name="Page" method="Post" action="SysUser_List.asp">
      <td height="50" align="center" valign="middle" bordercolor="#FFFFFF">共<font color="#FF2D00"><%=rs.RecordCount%></font>条&nbsp;&nbsp;<font color="#FF2D00"><%=Page%></font>/<font color="#FF2D00"><%=rs.pagecount%></font>&nbsp;&nbsp;
	  <a href="?Page=1&s_Level=<%=s_Level%>&s_User=<%=s_User%>&s_Name=<%=s_Name%>" class="Text_1">首页</a>
          <%If Page>1 Then%>
          <a href="?Page=<%=Page-1%>&s_Level=<%=s_Level%>&s_User=<%=s_User%>&s_Name=<%=s_Name%>" class="Text_1">上一页</a>
          <%else%>
        上一页
        <%End If%>
        <%If Page < rs.pagecount Then %>
        <a href="?Page=<%=Page+1%>&s_Level=<%=s_Level%>&s_User=<%=s_User%>&s_Name=<%=s_Name%>" class="Text_1">下一页</a>
        <%else%>
        下一页
        <%End If%>
        <a href="?Page=<%=rs.pagecount%>&s_Level=<%=s_Level%>&s_User=<%=s_User%>&s_Name=<%=s_Name%>" class="Text_1">尾页</a>
        <input name="Page" type="text" class="ipt2" id="Page" value="<%=Page%>" size="3">
        <input name="Submit2" type="submit" class="ipt2" value="GO">
        <input name="s_Level" type="hidden" value="<%=s_Level%>">
        <input name="s_User" type="hidden" value="<%=s_User%>">
        <input name="s_Name" type="hidden" value="<%=s_Name%>">
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