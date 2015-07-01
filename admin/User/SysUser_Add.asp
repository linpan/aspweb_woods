<!--#include file="../../Include/Class_Conn.asp"-->
<!--#include file="../../Include/Class_Main.asp"-->
<!--#include file="../../Include/Class_MD5.asp"-->
<% Super=1 '只允许超级管理员访问该页 %>
<!--#include file="../p8_Check.asp"-->
<%
	If Request("Submit") <> "" Then
		Dim s_User,s_Pass,s_Name,s_Level
		s_User   = Trim(Request("s_User"))
		s_Pass   = Trim(Request("s_Pass"))
		s_Name   = Trim(Request("s_Name"))
		s_Level  = Trim(Request("s_Level"))
		
		If s_User = "" Then
			Response.Write "<script>alert(""请填写用户名"");window.history.back();</script>"
			Response.End()
		End If
		
		If s_Pass = "" Then
			Response.Write "<script>alert(""请填写密码"");window.history.back();</script>"
			Response.End()
		End If
		
		If s_Name = "" Then
			Response.Write "<script>alert(""请填写姓名"");window.history.back();</script>"
			Response.End()
		End If
		
		If s_Level = "" Then
			Response.Write "<script>alert(""请填选择权限"");window.history.back();</script>"
			Response.End()
		End If
		
		Set rs=Server.CreateObject("Adodb.Recordset")
		rs.Open "Select s_User,s_Pass,s_Name,s_Level From p8_Super Where s_User = '"& s_User &"'",Conn,1,3
		
		If Not rs.Eof Then
			Response.Write "<script>alert(""用户名已存在，请更换其他用户名"");window.history.back();</script>"
			Response.End()
		Else
			rs.AddNew
			rs("s_User")  = s_User
			rs("s_Pass")  = MD5(s_Pass)
			rs("s_Name")  = s_Name
			rs("s_Level") = s_Level
			rs.Update
		End If

		rs.Close
		Set rs=Nothing
		
		CloseConn
		Response.Redirect "SysUser_List.asp?Tip=添加成功！"
		Response.End()
		
	End If
	
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>添加管理员</title>
<script type="text/javascript">top.window.aTitle.innerText='添加管理员'</script>
<meta http-equiv="pragma" content="no-cache" />
<meta http-equiv="Cache-Control" content="no-cache, must-revalidate" />
<meta http-equiv="expires" content="Wed, 26 Feb 1997 08:21:57 GMT" />
<meta http-equiv="expires" content="0" />
<link href="../css/Public.css" rel="stylesheet" type="text/css" /> 
<script type="text/javascript" src="../Include/Pub.js"></script>
<script type="text/javascript" src="../Include/TipBox.js"></script>
<script type="text/javascript">
function $p8(Obj){return document.getElementById(Obj);}
function Trim(strSource){return strSource.replace(/^\s*/,'').replace(/\s*$p8/,'');}
function CheckForm(){
	if(Trim($p8("s_User").value) == ''){
		alert("请填写用户名");
		$p8("s_User").focus();
		return false;
	}
	if(Trim($p8("s_Pass").value) == ''){
		alert("请填写密码");
		$p8("s_Pass").focus();
		return false;
	}
	if(Trim($p8("s_Pass").value) != Trim($p8("s_RePass").value)){
		alert("重复密码必须与密码相同");
		$p8("s_RePass").focus();
		return false;
	}
	if(Trim($p8("s_Name").value) == ''){
		alert("请填写姓名");
		$p8("s_Name").focus();
		return false;
	}

	var a = document.getElementsByName("s_Level");
	var num=0;
	for (var i=0; i<a.length; i++){
		if(a[i].checked) {
			num++;
		}
	}
	if(num==0) {
		alert("请选择权限");
		return false;
	}
	
	return true;
}
</script>
</head>

<body>
<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
<form name="AddForm" method="post" action="SysUser_Add.asp" onSubmit="return CheckForm()">
  <tr>
    <td width="80" height="30" align="right" bgcolor="#F8FBFE">用户名：</td>
    <td bgcolor="#F8FBFE"><input name="s_User" type="text" class="ipt3" id="s_User" maxlength="50" style="width:200px;"></td>
  </tr>
  <tr>
    <td height="30" align="right" bgcolor="#F8FBFE">密码：</td>
    <td bgcolor="#F8FBFE"><input name="s_Pass" type="password" class="ipt3" id="s_Pass" maxlength="50" style="width:200px;"></td>
  </tr>
  <tr>
    <td height="30" align="right" bgcolor="#F8FBFE">重复密码：</td>
    <td bgcolor="#F8FBFE"><input name="s_RePass" type="password" class="ipt3" id="s_RePass" maxlength="50" style="width:200px;"></td>
  </tr>
  <tr>
    <td height="30" align="right" bgcolor="#F8FBFE">姓名：</td>
    <td bgcolor="#F8FBFE"><input name="s_Name" type="text" class="ipt3" id="s_Name" maxlength="10" style="width:200px;"></td>
  </tr>
  <tr>
    <td height="30" align="right" bgcolor="#F8FBFE">权限：</td>
    <td bgcolor="#F8FBFE"><input type="radio" name="s_Level" value="1">超级管理员
      <input type="radio" name="s_Level" value="2">录入员</td>
  </tr>
  <tr>
    <td height="30" align="right" bgcolor="#F8FBFE">&nbsp;</td>
    <td bgcolor="#F8FBFE"><span style="padding:20px 0;">
      <input name="Submit" type="submit" class="btn2" value=" 添加 " >
    </span></td>
  </tr>
  <tr>
    <td height="30" align="right" bgcolor="#F8FBFE">&nbsp;</td>
    <td bgcolor="#F8FBFE">&nbsp;</td>
  </tr>
</form>
</table>
</body>
</html>
<%
	CloseConn
%>