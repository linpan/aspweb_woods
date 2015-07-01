<!--#include file="../../Include/Class_Conn.asp"-->
<!--#include file="../../Include/Class_Main.asp"-->
<!--#include file="../p8_Check.asp"-->
<%
'删除============================================================================================
If Request.QueryString("ac")="del" Then
	Dim FSO,i
	id          = Request("id")
	Page        = Request("Page")
	ClassID    = Request("ClassID")
	Title      = Request("Title")
	Source     = Request("Source")
	AddDate    = Request("AddDate")
	News_px    = Request("News_px")
	
	Set rs = Server.Createobject("Adodb.RecordSet")
	rs.open "Select Pic,SmallPic From p8_News Where id= " & id ,Conn,1,3
	
	Pic      = rs("Pic")
	SmallPic = rs("SmallPic")
	
	Set FSO = CreateObject(FsoName)
	If Instr(Pic,"|") Then
		Pic = Split(Pic,"|")
		For i=0 To Ubound(Pic)
			If Pic(i)<>"" Then
				If  FSO.FileExists(Server.MapPath(Pic(i))) Then
					FSO.Deletefile(server.MapPath(Pic(i)))
				End IF
			End If
		Next
	ElseIf Pic<>"" Then
		If FSO.FileExists(Server.MapPath(Pic)) Then
			FSO.Deletefile(server.MapPath(Pic))
		End IF
	End If
	
	If SmallPic<>"" Then
		If FSO.FileExists(Server.MapPath(SmallPic)) Then
			FSO.Deletefile(server.MapPath(SmallPic))
		End IF
	End If
	
	rs.Delete
	rs.Close
	Set rs = Nothing
	Set FSO = Nothing
	Response.Redirect "News_List.asp?Tip=删除成功！&Page="& Page &"&ClassID="& ClassID &"&Title="& Title &"&Source="& Source &"&AddDate="& AddDate &"&News_px="& News_px
	Response.End()
End If
'/删除============================================================================================

	Dim rs,Page,ClassID,Title,Source,AddDate,News_px,ClassID_Sql,Title_Sql,Source_Sql,AddDate_Sql,News_px_Sql
	ClassID    = Request("ClassID")
	Title      = Request("Title")
	Source     = Request("Source")
	AddDate    = Request("AddDate")
	News_px    = Request("News_px")
	
	If ClassID <>"" Then
		ClassID_Sql = " And (BigClass = "& ClassID &" Or SmallClass = "& ClassID &")"
	End If

	If Title <>"" Then
		Title_Sql = " And Title Like '%"& Title &"%' "
	End If
	
	If Source <>"" Then
		Source_Sql = " And Source Like '%"& Source &"%' "
	End If
	
	If AddDate <>"" Then
		AddDate_Sql = " And AddDate = '"& AddDate &"' "
	End If
	
	If News_px <>"" Then
		Select Case News_px
			Case "1" : News_px_Sql = " Order By AddDate Desc"
			Case "2" : News_px_Sql = " Order By AddDate Asc"
			Case "3" : News_px_Sql = " Order By Hits Desc"
			Case "4" : News_px_Sql = " Order By Hits Asc"
		End Select
	Else
		News_px_Sql = " Order By AddDate Desc"
	End If
	
	Set rs = Server.CreateObject("ADODB.RecordSet")
	rs.open "Select * From p8_News Where 1=1 "& ClassID_Sql & Title_Sql & Source_Sql & AddDate_Sql & News_px_Sql,Conn,1,1
	
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
<title>管理文章</title>
<script type="text/javascript">top.window.aTitle.innerText='管理文章'</script>
<meta http-equiv="pragma" content="no-cache" />
<meta http-equiv="Cache-Control" content="no-cache, must-revalidate" />
<meta http-equiv="expires" content="Wed, 26 Feb 1997 08:21:57 GMT" />
<meta http-equiv="expires" content="0" />
<link href="../css/Public.css" rel="stylesheet" type="text/css" /> 
</head>
<script type="text/javascript" src="../Include/TipBox.js"></script>
<script type="text/javascript" src="../Include/calendar.js"></script>
<script type="text/javascript" src="../Include/Pub.js"></script>
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
	<form name="form1" method="post" action="News_List.asp">
	
      <table style="margin-left:10px;" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="130">
		    <select name="ClassID" id="ClassID" style="width:120px;" class="ipt5">
		  <%
			Set rs2 = Server.Createobject("Adodb.RecordSet")
			rs2.open "Select id,ClassName From p8_Class Where ClassType='文章' And ClassLevel=1 Order By id Desc",Conn,1,1
			
				Do While Not rs2.Eof 
					
					If Clng(ClassID) = Clng(rs2("id")) Then
						Response.Write "<option value="""& rs2("id") &""" selected=""selected"">"& rs2("ClassName") &"</option>"
					Else
						Response.Write "<option value="""& rs2("id") &""">"& rs2("ClassName") &"</option>"
					End If
					
					Set rsnxt = Server.CreateObject("ADODB.RecordSet")
					rsnxt.open "Select id,ClassName From p8_Class Where ClassType='文章' And ClassLevel=2 And ParentID="& rs2("id") &" Order By id Desc",Conn,1,1
					
					Do While Not rsnxt.Eof 						
						If Clng(ClassID) = Clng(rsnxt("id")) Then
							Response.Write "<option value="""& rsnxt("id") &""" selected=""selected"">├ "& rsnxt("ClassName") &"</option>"
						Else
							Response.Write "<option value="""& rsnxt("id") &""">├ "& rsnxt("ClassName") &"</option>"
						End If
						rsnxt.MoveNext         
					Loop 
					
					rsnxt.Close
					Set rsnxt = Nothing

				
				rs2.MoveNext         
				Loop 

			rs2.Close
			Set rs2 = Nothing
		  %>
		</select>		</td>
          <td width="150">标题：<input style="width:100px;" name="Title" type="text" <%If Title="" Then Response.Write "class=""ipt2""" :Else Response.Write "class=""ipt"""%> value="<%=Title%>" maxlength="50"></td>
          <td width="140">来源：
            <input style="width:80px;" name="Source" type="text" <%If Source="" Then Response.Write "class=""ipt2""" :Else Response.Write "class=""ipt"""%> value="<%=Source%>" maxlength="50"></td>
          <td width="140">日期：
            <input name="AddDate" type="text" id="AddDate" style="width:80px;" value="<%=AddDate%>" maxlength="10" <%If AddDate="" Then Response.Write "class=""ipt2""" :Else Response.Write "class=""ipt"""%> onFocus="setday(this)" readonly="readonly"></td>
          <td width="100">
		    <select name="News_px" id="News_px" <%If News_px="" Then Response.Write "class=""ipt2""" :Else Response.Write "class=""ipt"""%>>
            <option value="">最近添加</option>
            <option value="1" <%If News_px = "1" Then Response.Write "selected"%>>最近添加</option>
            <option value="2" <%If News_px = "2" Then Response.Write "selected"%>>最早添加</option>
            <option value="3" <%If News_px = "3" Then Response.Write "selected"%>>人气最高</option>
            <option value="4" <%If News_px = "4" Then Response.Write "selected"%>>人气最低</option>
          </select></td>
          <td><input name="Submit" type="submit" class="btn1" value="搜索">
            &nbsp;&nbsp;<input name="Submit" type="button" class="btn1" onClick="window.location.href='News_List.asp?ClassID=<%=ClassID%>'" value="全部"></td>
          </tr>
      </table>
    </form>    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="1" cellpadding="0">
  <tr bgcolor="#E4EDF9">
    <td width="8%" height="25" align="center">编号</td>
    <td width="17%" align="center">分类</td>
	<td width="27%" align="center">标题</td>
    <td width="16%" align="center">日期</td>
    <td width="9%" align="center">点击</td>
    <td width="9%" align="center">编辑</td>
    <td width="14%" align="center">操作</td>
  </tr>
<%
If rs.RecordCount = 0 Then
	Response.Write "<tr bgcolor=""#F8FBFE""><td height=""400"" colspan=""8"" align=""center"">没有找到相关信息！</td></tr>"
Else
	Do While Not rs.Eof And Sum>0 
%>
	<tr onMouseOver="this.style.backgroundColor='#e7ffa6'" onMouseOut="this.style.backgroundColor='#F8FBFE'" bgcolor="#F8FBFE">
	<td height="25" align="center"><%=rs("id")%></td>
	<td align="center">
	  <%
		Set rs2 = Server.Createobject("Adodb.RecordSet")
		rs2.open "Select ClassName From p8_Class Where id = "& rs("BigClass") &"",Conn,1,1
		
			If Not rs2.Eof Then
				Response.Write rs2("ClassName")
			End If

		rs2.Close
		Set rs2 = Nothing
		
		If rs("SmallClass") <> "" Then
			Set rs2 = Server.Createobject("Adodb.RecordSet")
			rs2.open "Select ClassName From p8_Class Where id = "& rs("SmallClass") &"",Conn,1,1
			
				If Not rs2.Eof Then
					Response.Write " - " & rs2("ClassName")
				End If
	
			rs2.Close
			Set rs2 = Nothing
		End If
	  %>
	</td>
	<td style="padding-left:10px;"><a href="News_Edit.asp?id=<%=rs("id")%>&Page=<%=Page%>&ClassID=<%=ClassID%>&Title=<%=Title%>&Source=<%=Source%>&AddDate=<%=AddDate%>&News_px=<%=News_px%>">
	<%
		If rs("TitleColor")<>"" Then
			Response.Write "<font color="""& rs("TitleColor") &""">"& rs("Title") &"</font>"
		Else
			Response.Write rs("Title")
		End If
		
		If rs("SmallPic")<>"" Then Response.Write "<font class=""cGreen"">[图]</font>"
	%>
	</a></td>
	<td align="center"><%=cTime(rs("AddDate"),2)%></td>
	<td align="center"><%=rs("Hits")%></td>
	<td align="center"><%=rs("Admin")%></td>
	<td align="center">
	 <a href="News_Edit.asp?id=<%=rs("id")%>&Page=<%=Page%>&ClassID=<%=ClassID%>&Title=<%=Title%>&Source=<%=Source%>&AddDate=<%=AddDate%>&News_px=<%=News_px%>">修改</a> <a href="javascript:if(confirm('删除后不可恢复，是否继续？'))window.location.href='?ac=del&id=<%=rs("id")%>&Page=<%=Page%>&ClassID=<%=ClassID%>&Title=<%=Title%>&Source=<%=Source%>&AddDate=<%=AddDate%>&News_px=<%=News_px%>';">删除</a></td>
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
    <form name="Page" method="Post" action="News_List.asp">
      <td height="50" align="center" valign="middle" bordercolor="#FFFFFF">共<font color="#FF2D00"><%=rs.RecordCount%></font>条&nbsp;&nbsp;<font color="#FF2D00"><%=Page%></font>/<font color="#FF2D00"><%=rs.pagecount%></font>&nbsp;&nbsp;
	  <a href="?Page=1&ClassID=<%=ClassID%>&Title=<%=Title%>&Source=<%=Source%>&AddDate=<%=AddDate%>&News_px=<%=News_px%>" class="Text_1">首页</a>
          <%If Page>1 Then%>
          <a href="?Page=<%=Page-1%>&ClassID=<%=ClassID%>&Title=<%=Title%>&Source=<%=Source%>&AddDate=<%=AddDate%>&News_px=<%=News_px%>" class="Text_1">上一页</a>
          <%else%>
        上一页
        <%End If%>
        <%If Page < rs.pagecount Then %>
        <a href="?Page=<%=Page+1%>&ClassID=<%=ClassID%>&Title=<%=Title%>&Source=<%=Source%>&AddDate=<%=AddDate%>&News_px=<%=News_px%>" class="Text_1">下一页</a>
        <%else%>
        下一页
        <%End If%>
        <a href="?Page=<%=rs.pagecount%>&ClassID=<%=ClassID%>&Title=<%=Title%>&Source=<%=Source%>&AddDate=<%=AddDate%>&News_px=<%=News_px%>" class="Text_1">尾页</a>
        <input name="Page" type="text" class="ipt2" id="Page" value="<%=Page%>" size="3">
        <input name="Submit2" type="submit" class="ipt2" value="GO">
        <input name="ClassID" type="hidden" value="<%=ClassID%>">
        <input name="Title" type="hidden" value="<%=Title%>">
        <input name="Source" type="hidden" value="<%=Source%>">
        <input name="AddDate" type="hidden" value="<%=AddDate%>">
		<input name="News_px" type="hidden" value="<%=News_px%>">
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