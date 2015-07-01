<%
	Dim ExitPath
	ExitPath = 1
%>
<!--#include file="../Include/Class_Conn.asp"-->
<!--#include file="../Include/Class_Main.asp"-->
<!--#include file="../Include/Class_MD5.asp"-->
<!--#include file="p8_Check.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>后台管理中心</title>
<link href="css/Public.css?admin" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
	body {margin:0; overflow:hidden; background:url(images/Bg.jpg) 0 75px #dfe8f6 repeat-x;}
	#TopFrame {width:100%; height:28px;}
	#LeftFrame {width:136px; height:100%;}
	#main {width:100%; height:100%;}
-->
</style>
</head>
<body scroll="no">
<table border="0" cellpadding="0" cellspacing="0" height="100%" width="100%">
	<tr>
		<td colspan="2" height="27">
		<iframe frameborder="0" id="TopFrame" name="TopFrame" src="p8_Head.asp" scrolling="no"></iframe>
		</td>
	</tr>
    <tr>
		<td valign="top" id="tLeftFrame" style="width:136px; padding:5px; padding-right:0;">
		<iframe frameborder="0" id="LeftFrame" name="LeftFrame" src="p8_Left.asp"></iframe>
	  <td style="padding:5px; padding-left:0;">
	<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" style="border:1px solid #8db2e3;">
  <tr>
    <td height="23" background="images/Ti_bg.gif" style="font-size:13px; padding-top:3px; color:#30599c; font-weight:bold; padding-left:16px;"><div id="aTitle"></div></td>
  </tr>
  <tr>
    <td valign="top"><iframe frameborder="0" id="main" name="main" cols="136,*" src="System/Main.asp" style="height:100%;" scrolling="yes"></iframe></td>
  </tr>
</table>
	</td>
	</tr>
</table>
</body>
</html>