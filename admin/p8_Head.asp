<%
	Dim ExitPath
	ExitPath = 1
%>
<!--#include file="../Include/Class_Conn.asp"-->
<!--#include file="../Include/Class_Main.asp"-->
<!--#include file="../Include/Class_MD5.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta http-equiv="pragma" content="no-cache" />
<meta http-equiv="Cache-Control" content="no-cache, must-revalidate" />
<meta http-equiv="expires" content="Wed, 26 Feb 1997 08:21:57 GMT" />
<meta http-equiv="expires" content="0" />
<title>��̨��������</title>
<link href="css/Public.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="Include/Pub.js"></script>
<style type="text/css">
<!--
	.Headline {color:#cfddf2;}
-->
</style>
<style type="text/css">
	.Hd_Mu1 {padding:0px 10px 0 0; height:20px; line-height:20px; overflow:hidden; float:right;}
	.Hd_Mu1_txt {text-align:center; padding:0 8px; float:left;}
	.Hd_Mu1_out,.Hd_Mu1_pass {width:75px; text-align:center; float:left;}
	.Hd_Mu1_li {width:6px; background:url(images/Li_1.gif) no-repeat; float:left; overflow:hidden;}
	.Hd_Mu1_pass {padding-left:8px; background:url(images/ico_1.gif) 3px 4px no-repeat;}
	.Hd_Mu1_out {padding-left:8px; background:url(images/ico_2.gif) 6px 5px no-repeat;}
</style>
</head>

<body>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
  	<td width="110" height="27" align="center" background="images/Bar_Bg.gif"><img src="images/logo.gif" width="86" height="16" /></td>
    <td height="27" background="images/Bar_Bg.gif" style="padding:3px 20px 0 0; text-align:right;">
		<div class="Hd_Mu1">
			<ul>
				<li class="Hd_Mu1_txt">����,<%=Request.Cookies("Admin")("s_User")%>(<%=Request.Cookies("Admin")("s_Name")%>)&nbsp;</li>
				<li class="Hd_Mu1_li">&nbsp;</li>
				<%If Request.Cookies("Admin")("s_Level")="1" Then%>
				<li class="Hd_Mu1_txt"><a href="System/Config.asp" target="main">��վ����</a></li>
				<li class="Hd_Mu1_li">&nbsp;</li>
				<%End If%>
				<li class="Hd_Mu1_txt"><a href="Include/Editer/i6@web-(a)/default.asp?u=upload.asp?id=8" target="_blank">�ϴ�����</a></li>
				<li class="Hd_Mu1_li">&nbsp;</li>
				<li class="Hd_Mu1_txt"><a href="User/User_List.asp" target="main">ע���Ա����</a></li>
				<li class="Hd_Mu1_li">&nbsp;</li>
				<%If Request.Cookies("Admin")("s_Level")="1" Then%>
				<li class="Hd_Mu1_txt"><a href="User/SysUser_List.asp" target="main">����Ա����</a>(<a href="User/SysUser_Add.asp" target="main">���</a>)</li>
				<li class="Hd_Mu1_li">&nbsp;</li>
				<li class="Hd_Mu1_txt"><a href="User/SysLog.asp" target="main">��̨��½��־</a></li>
				<li class="Hd_Mu1_li">&nbsp;</li>
				<%End If%>
				<li class="Hd_Mu1_pass"><a href="#" onclick="openw('User/EditPass.asp','name1',600,400)">�޸�����</a></li>
				<li class="Hd_Mu1_out"><a href="p8_Exit.asp" target="_top">�˳���½</a></li>
			</ul>
		</div>	</td>
  </tr>
</table>
<!--����,<%=Request.Cookies("Admin")("s_User")%>(<%=Request.Cookies("Admin")("s_Name")%>)&nbsp;<span class="Headline">|</span>&nbsp;<a href="Include/Editer/i6@web-(a)/default.asp?u=upload.asp?id=8" target="_blank">�ϴ�����</a>&nbsp;<span class="Headline">|</span>&nbsp;<a href="#" onclick="openw('Manage/EditPass.asp','name1',600,400)">�޸�����</a>&nbsp;<span class="Headline">|</span>&nbsp;<a href="p8_Exit.asp" target="_top">�˳���½</a>
--></body>
</html>

