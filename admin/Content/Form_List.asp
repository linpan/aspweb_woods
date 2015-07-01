<!--#include file="../../Include/Class_Conn.asp"-->
<!--#include file="../../Include/Class_Main.asp"-->
<!--#include file="../p8_Check.asp"-->
<%
'删除============================================================================================
If Request.QueryString("ac")="del" Then
	Dim FSO,i
	id      = Request("id")
	Page    = Request("Page")
	ClassID = Request("ClassID")
	
	Set rs = Server.Createobject("Adodb.RecordSet")
	rs.open "Select id From p8_Form Where id= " & id ,Conn,1,3
	
	rs.Delete
	rs.Close
	Set rs = Nothing
	Set FSO = Nothing
	Response.Redirect "Form_List.asp?Tip=删除成功！&Page="& Page &"&ClassID="& ClassID
	Response.End()
End If
'/删除============================================================================================


	Dim rs,Page,FieldContent,ClassID,ClassNum
	
	ClassID    = Request("ClassID")
	
	If ClassID <>"" Then
		ClassID_Sql = " And ClassID = " & ClassID
	End If
	
	Set rs = Server.CreateObject("ADODB.RecordSet")
	rs.open "Select Num From p8_Class Where id = " & ClassID ,Conn,1,1
	ClassNum = rs("Num")
	rs.Close
	Set rs = Nothing
	
	Set rs = Server.CreateObject("ADODB.RecordSet")
	rs.open "Select * From p8_Form Where 1=1 "& ClassID_Sql & " Order By AddDate Desc",Conn,1,1
	
	rs.PageSize = 6
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
<title>管理表单</title>
<script type="text/javascript">top.window.aTitle.innerText='管理表单'</script>
<meta http-equiv="pragma" content="no-cache" />
<meta http-equiv="Cache-Control" content="no-cache, must-revalidate" />
<meta http-equiv="expires" content="Wed, 26 Feb 1997 08:21:57 GMT" />
<meta http-equiv="expires" content="0" />
<link href="../css/Public.css" rel="stylesheet" type="text/css" /> 
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
<%
If rs.RecordCount = 0 Then
	Response.Write "<br /><br /><br />没有找到相关信息！<br /><br />"
Else
	n = 1
	Do While Not rs.Eof And Sum>0 
		FieldContent = rs("FieldContent")
%>

<table width="98%" border="0" cellpadding="8" cellspacing="1" style="margin:20px 0;">
  <tr>
    <td colspan="2" align="right" bgcolor="#E4EDF9"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="2%" align="center"><%=n%>.</td>
        <td width="98%" align="right"><a href="javascript:if(confirm('删除后不可恢复，是否继续？'))window.location.href='?ac=del&id=<%=rs("id")%>&Page=<%=Page%>&ClassID=<%=ClassID%>';">删除</a></td>
      </tr>
    </table></td>
  </tr>
<%
		Set rsField = Server.Createobject("Adodb.RecordSet") '读取该分类下的所有自定义字段
		rsField.open "Select FieldName,Variable From p8_Field Where ClassNum='"& ClassNum &"' Order By id Asc",Conn,1,1
		
		Do While Not rsField.Eof
			
%>
  <tr>
    <td width="100" bgcolor="#F8FBFE" align="right"><%=rsField("FieldName")%>：</td>
    <td bgcolor="#F8FBFE"><%=GetFormField(rsField("Variable"),rs("id"))%></td>
  </tr>
<%
			
		rsField.MoveNext     
		Loop 
%>
</table>
<%
	rs.MoveNext  
	n = n + 1   
	Sum = Sum - 1     
	Loop 
End If
%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <form name="Page" method="Post" action="Form_List.asp">
      <td height="50" align="center" valign="middle" bordercolor="#FFFFFF">共<font color="#FF2D00"><%=rs.RecordCount%></font>条&nbsp;&nbsp;<font color="#FF2D00"><%=Page%></font>/<font color="#FF2D00"><%=rs.pagecount%></font>&nbsp;&nbsp; <a href="?Page=1&ClassID=<%=ClassID%>" class="Text_1">首页</a>
          <%If Page>1 Then%>
          <a href="?Page=<%=Page-1%>&ClassID=<%=ClassID%>" class="Text_1">上一页</a>
          <%else%>
        上一页
        <%End If%>
        <%If Page < rs.pagecount Then %>
        <a href="?Page=<%=Page+1%>&ClassID=<%=ClassID%>" class="Text_1">下一页</a>
        <%else%>
        下一页
        <%End If%>
        <a href="?Page=<%=rs.pagecount%>&ClassID=<%=ClassID%>" class="Text_1">尾页</a>
        <input name="Page" type="text" class="ipt2" id="Page" value="<%=Page%>" size="3">
        <input name="Submit2" type="submit" class="ipt2" value="GO">
        <input name="ClassID" type="hidden" value="<%=ClassID%>">
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