<!--#include file="../../Include/Class_Conn.asp"-->
<!--#include file="../../Include/Class_MD5.asp"-->
<!--#include file="../p8_Check.asp"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title></title>
<link href="../css/Public.css?admin" rel="stylesheet" type="text/css" />
<meta http-equiv="pragma" content="no-cache" />
<meta http-equiv="Cache-Control" content="no-cache, must-revalidate" />
<meta http-equiv="expires" content="Wed, 26 Feb 1997 08:21:57 GMT" />
<meta http-equiv="expires" content="0" />
<style type="text/css">
<!--
	#Menu td {padding-top:4px; cursor:pointer; padding-left:20px;}
	#Menu td a:link {text-decoration:none;}
	#Menu td a:hover {text-decoration:underline;}
	.Btn {background:url(../images/ico_3.gif) 8px 9px no-repeat;}
	.Btn_over {background:url(../images/ico_3.gif) #eaf3fd 8px 9px no-repeat;}
	/*.Btn_over a:link {font-weight:bold;}*/
-->
</style>
</head>

<body>
<table width="136" height="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td height="23" align="center" background="../images/Ti_bg.gif" style="border:1px solid #8db2e3; border-bottom:none; font-size:13px; padding-top:3px;"><strong>��Ա����</strong></td>
    <td width="6" rowspan="2" style="background:url(../images/Bg.jpg) 0 -5px #dfe8f6 repeat-x;"></td>
  </tr>
  <tr>
    <td width="130" valign="top" style="border:1px solid #8db2e3;"><table id="Menu" width="100%" border="0" cellpadding="0" cellspacing="0">
	<%If Request.Cookies("Admin")("s_Level")="1" Then%>
	<tr>
	  <td onMouseOver="this.className='Btn_over';" onMouseOut="this.className='Btn';" class="Btn" height="25"> <a href="SysUser_List.asp" target="main">��վ����Ա</a> | <a href="SysUser_Add.asp" target="main">���</a></td>
	</tr>
        <tr>
          <td onMouseOver="this.className='Btn_over';" onMouseOut="this.className='Btn';" class="Btn" height="25"><a href="SysLog.asp" target="main">��̨��¼��־</a></td>
        </tr>
	<%End If%>
        <tr>
          <td onMouseOver="this.className='Btn_over';" onMouseOut="this.className='Btn';" class="Btn" height="25"><a href="User_List.asp" target="main">ע���Ա</a></td>
        </tr>
		<tr>
          <td height="1" bgcolor="#DFE8F7"></td>
        </tr>

      </table></td>
  </tr>
</table>
</body>
</html>

