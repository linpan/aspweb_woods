<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!--#include file="Include/Class_Conn.asp"-->
<!--#include file="Include/Class_Main.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>会员登录</title>
<link href="css/login.css" rel="stylesheet" type="text/css" />
<!--[if IE 6]>
	<script type="text/javascript" src="js/DD_belatedPNG.js"></script>
	
	<script type="text/javascript">
	 DD_belatedPNG.fix('img');
	</script>
<![endif]-->
</head>

<body>
<div id="bg">
	<div id="logo"><img src="images/login_logo.png" width="139" height="32" /></div>
    <div id="kh"><img src="images/login_img.png" width="321" height="34" /></div>
  <div id="login"><%If Session("UserName")<>"" Then%>
<table width="100%" border="0" cellspacing="5" cellpadding="0">
		<tr>
		  <td colspan="2">您好，<%=Session("UserName")%><br />
	      <a href="<%=SiteDir%>User/Main.asp">修改资料</a>?|?<a href="<%=SiteDir%>User/LoginOut.asp">退出登录</a></td>
	  </tr>
</table>
<%Else%>
<table width="100%" border="0" cellspacing="5" cellpadding="0">
	<form method="post" action="<%=SiteDir%>User/Login.asp">
		<tr>
		  <td width="56" height="20" align="right">用户名：</td>
	    <td width="101"><input name="UserName" type="text" class="ipt1" id="UserName" value="" maxlength="50" ></td></tr>
		<tr>
		  <td height="20" align="right">密码：</td>
		  <td><input name="PassWord" type="password" class="ipt1" id="PassWord" value="" maxlength="50"></td>
		</tr>
		<tr>
		  <td height="20" colspan="2" align="center"><a href="<%=SiteDir%>User/Register.asp">立即注册</a>  <a href="<%=SiteDir%>User/GetPass.asp">找回密码</a></td>
	    </tr>
      <tr>
		  <td height="35" align="right" valign="bottom"></td>
		  <td valign="bottom" id="annu"><input name="Submit" type="submit" class="btn1" value="" id="dl" /><input name="" type="reset" id="cz" value="" /></td>
	  </tr>
      </form>
      </table></div>
	
<%End If%>
</div>
</body>
</html>
