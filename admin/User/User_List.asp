<!--#include file="../../Include/Class_Conn.asp"-->
<!--#include file="../../Include/Class_Main.asp"-->
<!--#include file="../p8_Check.asp"-->
<%
'ɾ��============================================================================================
If Request.QueryString("ac")="del" Then
	id       = Request("id")
	Page     = Request("Page")
	UserName = Request("UserName")
	FullName = Request("FullName")
	Company  = Request("Company")
	Tel      = Request("Tel")
	Mob      = Request("Mob")
	
	Set rs = Server.Createobject("Adodb.RecordSet")
	rs.open "Select id From p8_User Where id= " & id ,Conn,1,3
	
	rs.Delete
	rs.Close
	Set rs = Nothing
	Response.Redirect "User_List.asp?Tip=ɾ���ɹ���&Page="& Page &"&UserName="& UserName &"&FullName="& FullName &"&Company="& Company &"&Tel="& Tel &"&Mob="& Mob
	Response.End()
End If
'/ɾ��============================================================================================

'�������============================================================================================
	If Request("SetUserState")<>"" then
		id            = Trim(Request("id"))
		SetUserState  = Trim(Request("SetUserState"))

		conn.execute "UpDate p8_User Set UserState="& SetUserState &" Where id="& id
		
		Tip = "���óɹ���"
	End If
'/�������============================================================================================

	Dim rs,Page,UserName,FullName,Company,Tel,Mob,UserName_Sql,FullName_Sql,Company_Sql,Tel_Sql,Mob_Sql
	UserName = Request("UserName")
	FullName = Request("FullName")
	Company  = Request("Company")
	Tel      = Request("Tel")
	Mob      = Request("Mob")
	
	If UserName <>"" Then
		UserName_Sql = " And UserName Like '%"& UserName &"%' "
	End If

	If FullName <>"" Then
		FullName_Sql = " And FullName Like '%"& FullName &"%' "
	End If
	
	If Company <>"" Then
		Company_Sql = " And Company Like '%"& Company &"%' "
	End If
	
	If Tel <>"" Then
		Tel_Sql = " And Tel Like '%"& Tel &"%' "
	End If
	
	If Mob <>"" Then
		Mob_Sql = " And Mob Like '%"& Mob &"%' "
	End If
	
	Set rs = Server.CreateObject("ADODB.RecordSet")
	rs.open "Select * From p8_User Where 1=1 "& UserName_Sql & FullName_Sql & Company_Sql & Tel_Sql & Mob_Sql &" Order By id Desc",Conn,1,1
	
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
<title>ע���Ա</title>
<script type="text/javascript">top.window.aTitle.innerText='ע���Ա'</script>
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
    <td height="30" bgcolor="#eaf3fd" style="border-bottom:1px solid #b5cef0;">
	<form name="form1" method="post" action="User_List.asp">
	
      <table style="margin-left:10px;" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="140">�û�����<input style="width:80px;" name="UserName" type="text" <%If UserName="" Then Response.Write "class=""ipt2""" :Else Response.Write "class=""ipt"""%> value="<%=UserName%>" maxlength="50"></td>
          <td width="140">������
            <input style="width:80px;" name="FullName" type="text" <%If FullName="" Then Response.Write "class=""ipt2""" :Else Response.Write "class=""ipt"""%> value="<%=FullName%>" maxlength="50"></td>
          <td width="140">�绰��
            <input style="width:80px;" name="Tel" type="text" <%If Tel="" Then Response.Write "class=""ipt2""" :Else Response.Write "class=""ipt"""%> value="<%=Tel%>" maxlength="50"></td>
          <td width="140">�ֻ���
            <input style="width:80px;" name="Mob" type="text" <%If Mob="" Then Response.Write "class=""ipt2""" :Else Response.Write "class=""ipt"""%> value="<%=Mob%>" maxlength="50"></td>
          <td width="160">��˾���ƣ�
            <input style="width:80px;" name="Company" type="text" <%If Company="" Then Response.Write "class=""ipt2""" :Else Response.Write "class=""ipt"""%> value="<%=Company%>" maxlength="50"></td>
          <td><input name="Submit" type="submit" class="btn1" value="����">
            &nbsp;&nbsp;<input name="Submit" type="button" class="btn1" onClick="window.location.href='User_List.asp'" value="ȫ��"></td>
          </tr>
      </table>
    </form>    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="1" cellpadding="0">
  <tr bgcolor="#E4EDF9">
    <td width="12%" height="25" align="center">�û���</td>
    <td width="7%" align="center">����</td>
	<td width="13%" align="center">�绰</td>
    <td width="13%" align="center">�ֻ�</td>
    <td width="17%" align="center">��˾����</td>
    <td width="17%" align="center">ע��ʱ��</td>
    <td width="7%" align="center">���</td>
    <td width="14%" align="center">����</td>
  </tr>
<%
If rs.RecordCount = 0 Then
	Response.Write "<tr bgcolor=""#F8FBFE""><td height=""400"" colspan=""8"" align=""center"">û���ҵ������Ϣ��</td></tr>"
Else
	Do While Not rs.Eof And Sum>0 
%>
	<tr onMouseOver="this.style.backgroundColor='#e7ffa6'" onMouseOut="this.style.backgroundColor='#F8FBFE'" bgcolor="#F8FBFE">
	<td height="25" align="center"><a href="User_Session.asp?UserName=<%=rs("UserName")%>" target="_blank"><%=rs("UserName")%></a></td>
	<td align="center"><%=rs("FullName")%></td>
	<td align="center"><%=rs("Tel")%></td>
	<td align="center"><%=rs("Mob")%></td>
	<td align="center"><%=rs("Company")%></td>
	<td align="center"><%=rs("AddDate")%></td>
	<td align="center">
<%
	If rs("UserState") = 1 Then
%>
	<a href="?SetUserState=0&id=<%=rs("id")%>&Page=<%=Page%>&UserName=<%=UserName%>&FullName=<%=FullName%>&Tel=<%=Tel%>&Mob=<%=Mob%>&Company=<%=Company%>">��ͨ��</a>
<%
	Else
%>
	<a href="?SetUserState=1&id=<%=rs("id")%>&Page=<%=Page%>&UserName=<%=UserName%>&FullName=<%=FullName%>&Tel=<%=Tel%>&Mob=<%=Mob%>&Company=<%=Company%>"><font color="#FF0000">�����</font></a>
<%
	End If
%>
	</td>
	<td align="center">
	 <a href="User_Session.asp?UserName=<%=rs("UserName")%>" target="_blank">����</a> <a href="javascript:if(confirm('ɾ���󲻿ɻָ����Ƿ������'))window.location.href='?ac=del&id=<%=rs("id")%>&Page=<%=Page%>&UserName=<%=UserName%>&FullName=<%=FullName%>&Tel=<%=Tel%>&Mob=<%=Mob%>&Company=<%=Company%>';">ɾ��</a></td>
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
    <form name="Page" method="Post" action="User_List.asp">
      <td height="50" align="center" valign="middle" bordercolor="#FFFFFF">��<font color="#FF2D00"><%=rs.RecordCount%></font>��&nbsp;&nbsp;<font color="#FF2D00"><%=Page%></font>/<font color="#FF2D00"><%=rs.pagecount%></font>&nbsp;&nbsp;
	  <a href="?Page=1&UserName=<%=UserName%>&FullName=<%=FullName%>&Tel=<%=Tel%>&Mob=<%=Mob%>&Company=<%=Company%>" class="Text_1">��ҳ</a>
          <%If Page>1 Then%>
          <a href="?Page=<%=Page-1%>&UserName=<%=UserName%>&FullName=<%=FullName%>&Tel=<%=Tel%>&Mob=<%=Mob%>&Company=<%=Company%>" class="Text_1">��һҳ</a>
          <%else%>
        ��һҳ
        <%End If%>
        <%If Page < rs.pagecount Then %>
        <a href="?Page=<%=Page+1%>&UserName=<%=UserName%>&FullName=<%=FullName%>&Tel=<%=Tel%>&Mob=<%=Mob%>&Company=<%=Company%>" class="Text_1">��һҳ</a>
        <%else%>
        ��һҳ
        <%End If%>
        <a href="?Page=<%=rs.pagecount%>&UserName=<%=UserName%>&FullName=<%=FullName%>&Tel=<%=Tel%>&Mob=<%=Mob%>&Company=<%=Company%>" class="Text_1">βҳ</a>
        <input name="Page" type="text" class="ipt2" id="Page" value="<%=Page%>" size="3">
        <input name="Submit2" type="submit" class="ipt2" value="GO">
        <input name="UserName" type="hidden" value="<%=UserName%>">
        <input name="FullName" type="hidden" value="<%=FullName%>">
        <input name="Tel" type="hidden" value="<%=Tel%>">
		<input name="Mob" type="hidden" value="<%=Mob%>">
		<input name="Company" type="hidden" value="<%=Company%>">
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