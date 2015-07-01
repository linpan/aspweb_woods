<!--#include file="../../Include/Class_Conn.asp"-->
<!--#include file="../../Include/Class_Main.asp"-->
<% Super=1 '只允许超级管理员访问该页 %>
<!--#include file="../p8_Check.asp"-->
<%
	If Request("Submit") <> "" Then
		SiteDir    = Trim(Request("SiteDir"))
		FsoName    = Trim(Request("FsoName"))
		UserCheck  = Trim(Request("UserCheck"))
		UserLgErr  = Trim(Request("UserLgErr"))
		UserLgLock = Trim(Request("UserLgLock"))
		SysLgErr   = Trim(Request("SysLgErr"))
		SysLgLock  = Trim(Request("SysLgLock"))
		SmtpEmail  = Trim(Request("SmtpEmail"))
		SmtpUser   = Trim(Request("SmtpUser"))
		SmtpPass   = Trim(Request("SmtpPass"))
		SmtpServer = Trim(Request("SmtpServer"))
		If UserCheck = "" Then UserCheck = 0
		
		If SiteDir = "" Then
			Response.Write "<script>alert(""请填写网站运行目录"");window.history.back();</script>"
			Response.End()
		End If
		
		Set rs=Server.CreateObject("Adodb.Recordset")
		rs.Open "Select Top 1 SiteDir,FsoName,UserCheck,UserLgErr,UserLgLock,SysLgErr,SysLgLock,SmtpEmail,SmtpUser,SmtpPass,SmtpServer From p8_Config Order By id Asc",Conn,1,3
		
		If rs.Eof Then
			rs.AddNew
		End If
		
		If SiteDir    <> ""  Then rs("SiteDir")    = SiteDir
		If FsoName    <> ""  Then rs("FsoName")    = FsoName
		If UserCheck  <> ""  Then rs("UserCheck")  = UserCheck
		If UserLgErr  <> "" And isNumeric(UserLgErr)  Then rs("UserLgErr")  = UserLgErr
		If UserLgLock <> "" And isNumeric(UserLgLock) Then rs("UserLgLock") = UserLgLock
		If SysLgErr   <> "" And isNumeric(SysLgErr)   Then rs("SysLgErr")   = SysLgErr
		If SysLgLock  <> "" And isNumeric(SysLgLock)  Then rs("SysLgLock")  = SysLgLock
		rs("SmtpEmail")  = SmtpEmail
		rs("SmtpUser")   = SmtpUser
		rs("SmtpPass")   = SmtpPass
		rs("SmtpServer") = SmtpServer

		rs.Update
		rs.Close
		Set rs=Nothing
		
		CloseConn
		Response.Redirect "Config.asp?Tip=设置成功！"
		Response.End()
		
	End If
	
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>网站参数</title>
<script type="text/javascript">top.window.aTitle.innerText='网站参数'</script>
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
	if(Trim($p8("SiteDir").value) == ''){
		alert("请填写网站运行目录");
		$p8("SiteDir").focus();
		return false;
	}
	if(Trim($p8("FsoName").value) == ''){
		alert("请填写服务器FSO组件名");
		$p8("FsoName").focus();
		return false;
	}
	if(Trim($p8("UserLgErr").value) == ''){
		alert("请填写会员登录允许错误次数");
		$p8("UserLgErr").focus();
		return false;
	}
	if(Trim($p8("UserLgLock").value) == ''){
		alert("请填写达到错误次数锁定时长");
		$p8("UserLgLock").focus();
		return false;
	}
	if(Trim($p8("SysLgErr").value) == ''){
		alert("请填写后台登录允许错误次数");
		$p8("SysLgErr").focus();
		return false;
	}
	if(Trim($p8("SysLgLock").value) == ''){
		alert("请填写达到错误次数锁定时长");
		$p8("SysLgLock").focus();
		return false;
	}
	return true;
}
</script>
</head>

<body>
<%
	Dim Tip
	Tip = Request.QueryString("Tip")
	If Tip <> "" Then
		Response.Write "<script type=""text/javascript"">window.onload=function(){new x.creat(1, 41, 5, 10, '"& Tip &"');}</script>"
	End If

	Set rs=Server.CreateObject("Adodb.Recordset")
	rs.Open "Select Top 1 SiteDir,FsoName,UserCheck,UserLgErr,UserLgLock,SysLgErr,SysLgLock,SmtpEmail,SmtpUser,SmtpPass,SmtpServer From p8_Config Order By id Asc",Conn,1,3
	
	If rs.Eof Then
		Response.Write "数据库系统参数丢失，请检查数据库！"
		Response.End()
	End If
%>
<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
<form method="post" action="Config.asp" onSubmit="return CheckForm()">
  <tr>
    <td width="150" height="30" align="right" bgcolor="#F8FBFE">网站运行目录：</td>
    <td bgcolor="#F8FBFE"><input name="SiteDir" type="text" class="ipt3" id="SiteDir" value="<%=rs("SiteDir")%>" maxlength="50" style="width:200px;">
      <span class="cGray">开始和结尾必须带/</span></td>
  </tr>
  <tr>
    <td height="30" align="right" bgcolor="#F8FBFE">服务器FSO组件名：</td>
    <td bgcolor="#F8FBFE"><input name="FsoName" type="text" class="ipt3" id="FsoName" value="<%=rs("FsoName")%>" maxlength="50" style="width:200px;"></td>
  </tr>
  <tr>
    <td height="30" align="right" bgcolor="#F8FBFE">会员注册审核：</td>
    <td bgcolor="#F8FBFE"><input name="UserCheck" type="checkbox" id="UserCheck" value="1" <%If rs("UserCheck")=1 Then Response.Write "checked=""checked"""%>>需要审核</td>
  </tr>
  <tr>
    <td height="30" align="right" bgcolor="#F8FBFE">会员登录允许错误次数：</td>
    <td bgcolor="#F8FBFE"><input name="UserLgErr" type="text" class="ipt3" id="UserLgErr" value="<%=rs("UserLgErr")%>" maxlength="6" style="width:200px;"></td>
  </tr>
  <tr>
    <td height="30" align="right" bgcolor="#F8FBFE">达到错误次数锁定时长：</td>
    <td bgcolor="#F8FBFE"><input name="UserLgLock" type="text" class="ipt3" id="UserLgLock" value="<%=rs("UserLgLock")%>" maxlength="6" style="width:200px;">
      分钟 <span class="cGray">当注册会员登录次数达到以上显示，则多少分钟内不允许登录</span> </td>
  </tr>
  <tr>
    <td height="30" align="right" bgcolor="#F8FBFE">后台登录允许错误次数：</td>
    <td bgcolor="#F8FBFE"><input name="SysLgErr" type="text" class="ipt3" id="SysLgErr" value="<%=rs("SysLgErr")%>" maxlength="6" style="width:200px;"></td>
  </tr>
  <tr>
    <td height="30" align="right" bgcolor="#F8FBFE">达到错误次数锁定时长：</td>
    <td bgcolor="#F8FBFE"><input name="SysLgLock" type="text" class="ipt3" id="SysLgLock" value="<%=rs("SysLgLock")%>" maxlength="6" style="width:200px;">
      分钟 <span class="cGray">当后台用户登录次数达到以上显示，则多少分钟内不允许登录</span> </td>
  </tr>
  <tr>
    <td height="30" align="right" bgcolor="#F8FBFE">SMTP邮件地址：</td>
    <td bgcolor="#F8FBFE"><input name="SmtpEmail" type="text" class="ipt3" id="SmtpEmail" value="<%=rs("SmtpEmail")%>" maxlength="50" style="width:200px;">
      &nbsp;<span class="cGray">用于向用户发送邮件，格式如：moumou@qq.com</span></td>
  </tr>
  <tr>
    <td height="30" align="right" bgcolor="#F8FBFE">SMTP邮件用户名：</td>
    <td bgcolor="#F8FBFE"><input name="SmtpUser" type="text" class="ipt3" id="SmtpUser" value="<%=rs("SmtpUser")%>" maxlength="50" style="width:200px;">
      &nbsp;<span class="cGray">格式如：moumou</span></td>
  </tr>
  <tr>
    <td height="30" align="right" bgcolor="#F8FBFE">SMTP邮件密码：</td>
    <td bgcolor="#F8FBFE"><input name="SmtpPass" type="text" class="ipt3" id="SmtpPass" value="<%=rs("SmtpPass")%>" maxlength="50" style="width:200px;"></td>
  </tr>
  <tr>
    <td height="30" align="right" bgcolor="#F8FBFE">SMTP服务器：</td>
    <td bgcolor="#F8FBFE"><input name="SmtpServer" type="text" class="ipt3" id="SmtpServer" value="<%=rs("SmtpServer")%>" maxlength="50" style="width:200px;">
      &nbsp;<span class="cGray">格式如：smtp.qq.com</span></td>
  </tr>
  <tr>
    <td height="30" align="right" bgcolor="#F8FBFE">&nbsp;</td>
    <td bgcolor="#F8FBFE"><span style="padding:20px 0;">
      <input name="Submit" type="submit" class="btn2" value=" 保存 " >
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
	rs.Close
	Set rs=Nothing
	CloseConn
%>