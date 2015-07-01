<!--#include file="../../Include/Class_Conn.asp"-->
<!--#include file="../../Include/Class_Main.asp"-->
<% Super=1 '只允许超级管理员访问该页 %>
<!--#include file="../p8_Check.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>登录框 - 数据调用</title>
<script type="text/javascript">top.window.aTitle.innerText='数据调用'</script>
<meta http-equiv="pragma" content="no-cache" />
<meta http-equiv="Cache-Control" content="no-cache, must-revalidate" />
<meta http-equiv="expires" content="Wed, 26 Feb 1997 08:21:57 GMT" />
<meta http-equiv="expires" content="0" />
<link href="../css/Public.css" rel="stylesheet" type="text/css" /> 
<script type="text/javascript" src="../Include/Pub.js"></script>
<script type="text/javascript" src="../Include/TipBox.js"></script>
<script type="text/javascript">
function Copy(Obj){ 
	var clipBoardContent = $Get(Obj).value; 
	$Get(Obj).select();
	window.clipboardData.setData("Text",clipBoardContent); 
	//alert("复制成功!"); 
	new x.creat(1, 41, 5, 10, '复制成功!');
} 
</script>
</head>

<body>
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
    <tr>
      <td bgcolor="#F8FBFE"><table border="0" cellspacing="10" cellpadding="0">
        <tr>
          <td width="80" height="30" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Code_News_List.asp';">文章列表</td>
          <td width="80" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Code_News_View.asp';">文章显示</td>
          <td width="80" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Code_Pic_List.asp';">图片列表</td>
          <td width="80" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Code_Pic_View.asp';">图片显示</td>
          <td width="80" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Code_Page.asp';">单页</td>
          <td width="80" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Code_User.asp';">表单</td>
          <td width="80" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Code_Service.asp';">在线客服</td>
          <td width="80" align="center" class="Tab1_over">登录框</td>
        </tr>
      </table>	  </td>
    </tr>
    <tr>
      <td bgcolor="#F8FBFE" style="padding:10px;"><table width="100%" border="0" cellpadding="10" cellspacing="1" bgcolor="#E4EDF9">
          <tr>
            <td bgcolor="#FFFFFF" style="line-height:160%;"><strong>调用说明：</strong><br>
            <span class="cGray">放置代码前，请保证需要放置代码的文件扩展名为.asp，如asp文件中包含该代码“&lt;%@LANGUAGE=&quot;VBSCRIPT&quot; CODEPAGE=&quot;936&quot;%&gt;”，请将其删除。</span></td>
          </tr>
      </table></td>
    </tr>
  </table>

<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
<tr>
  <td width="74" height="30" align="center" bgcolor="#F8FBFE">数据通讯：</td>
  <td bgcolor="#F8FBFE"><textarea id="Code1" class="ipt3" style="width:550px; height:60px;" readonly="readonly">&lt;!--#include file="Include/Class_Conn.asp"--&gt;
&lt;!--#include file="Include/Class_Main.asp"--&gt;
</textarea>
      <br>
      <input name="Submit" type="button" class="btn3" onClick="Copy('Code1')" value="复制以上代码">
    将以上代码放到网页最顶部（如页面中已存在，则不需要重复放置）</td>
</tr>
<tr>
  <td height="30" align="center" bgcolor="#F8FBFE">注册代码：</td>
  <td bgcolor="#F8FBFE">
<textarea id="Code2" class="ipt3" style="width:550px; height:280px;" readonly="readonly">
&lt;%If Session("UserName")&lt;>"" Then%>
&lt;table width="100%" border="0" cellspacing="5" cellpadding="0">
		&lt;tr>
		  &lt;td colspan="2">您好，&lt;%=Session("UserName")%>&lt;br />
	      &lt;a href="&lt;%=SiteDir%>User/Main.asp">修改资料&lt;/a>&nbsp;|&nbsp;&lt;a href="&lt;%=SiteDir%>User/LoginOut.asp">退出登录&lt;/a>&lt;/td>
		  &lt;/tr>
&lt;/table>
&lt;%Else%>
&lt;table width="100%" border="0" cellspacing="5" cellpadding="0">
	&lt;form method="post" action="&lt;%=SiteDir%>User/Login.asp">
		&lt;tr>
		  &lt;td width="100" height="30" align="right">用户名：&lt;/td>
		  &lt;td>&lt;input name="UserName" type="text" class="ipt1" id="UserName" value="" maxlength="50" >&lt;/td>&lt;/tr>
		&lt;tr>
		  &lt;td height="30" align="right">密码：&lt;/td>
		  &lt;td>&lt;input name="PassWord" type="password" class="ipt1" id="PassWord" value="" maxlength="50">&lt;/td>
		&lt;/tr>
		&lt;tr>
		  &lt;td height="30" align="right">&lt;/td>
		  &lt;td>&lt;input name="Submit" type="submit" class="btn1" value="登录" />&lt;/td>
	  &lt;/tr>
		&lt;tr>
		  &lt;td height="30" align="right">&lt;/td>
		  &lt;td>&lt;a href="&lt;%=SiteDir%>User/Register.asp">立即注册&lt;/a>  &lt;a href="&lt;%=SiteDir%>User/GetPass.asp">找回密码&lt;/a>&lt;/td>
	  &lt;/tr>
	&lt;/form>
&lt;/table>
&lt;%End If%>
</textarea>
      <br>
      <input name="Submit" type="button" class="btn3" onClick="Copy('Code2')" value="复制以上代码">
    将以上代码放到需要显示登录框的位置</td>
</tr>
<tr>
  <td height="30" align="center" bgcolor="#F8FBFE">&nbsp;</td>
  <td bgcolor="#F8FBFE">&nbsp;</td>
</tr>
</table>

</body>
</html>
<%
	CloseConn
%>