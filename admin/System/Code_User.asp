<!--#include file="../../Include/Class_Conn.asp"-->
<!--#include file="../../Include/Class_Main.asp"-->
<% Super=1 'ֻ����������Ա���ʸ�ҳ %>
<!--#include file="../p8_Check.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��¼�� - ���ݵ���</title>
<script type="text/javascript">top.window.aTitle.innerText='���ݵ���'</script>
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
	//alert("���Ƴɹ�!"); 
	new x.creat(1, 41, 5, 10, '���Ƴɹ�!');
} 
</script>
</head>

<body>
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
    <tr>
      <td bgcolor="#F8FBFE"><table border="0" cellspacing="10" cellpadding="0">
        <tr>
          <td width="80" height="30" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Code_News_List.asp';">�����б�</td>
          <td width="80" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Code_News_View.asp';">������ʾ</td>
          <td width="80" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Code_Pic_List.asp';">ͼƬ�б�</td>
          <td width="80" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Code_Pic_View.asp';">ͼƬ��ʾ</td>
          <td width="80" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Code_Page.asp';">��ҳ</td>
          <td width="80" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Code_User.asp';">��</td>
          <td width="80" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Code_Service.asp';">���߿ͷ�</td>
          <td width="80" align="center" class="Tab1_over">��¼��</td>
        </tr>
      </table>	  </td>
    </tr>
    <tr>
      <td bgcolor="#F8FBFE" style="padding:10px;"><table width="100%" border="0" cellpadding="10" cellspacing="1" bgcolor="#E4EDF9">
          <tr>
            <td bgcolor="#FFFFFF" style="line-height:160%;"><strong>����˵����</strong><br>
            <span class="cGray">���ô���ǰ���뱣֤��Ҫ���ô�����ļ���չ��Ϊ.asp����asp�ļ��а����ô��롰&lt;%@LANGUAGE=&quot;VBSCRIPT&quot; CODEPAGE=&quot;936&quot;%&gt;�����뽫��ɾ����</span></td>
          </tr>
      </table></td>
    </tr>
  </table>

<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
<tr>
  <td width="74" height="30" align="center" bgcolor="#F8FBFE">����ͨѶ��</td>
  <td bgcolor="#F8FBFE"><textarea id="Code1" class="ipt3" style="width:550px; height:60px;" readonly="readonly">&lt;!--#include file="Include/Class_Conn.asp"--&gt;
&lt;!--#include file="Include/Class_Main.asp"--&gt;
</textarea>
      <br>
      <input name="Submit" type="button" class="btn3" onClick="Copy('Code1')" value="�������ϴ���">
    �����ϴ���ŵ���ҳ�������ҳ�����Ѵ��ڣ�����Ҫ�ظ����ã�</td>
</tr>
<tr>
  <td height="30" align="center" bgcolor="#F8FBFE">ע����룺</td>
  <td bgcolor="#F8FBFE">
<textarea id="Code2" class="ipt3" style="width:550px; height:280px;" readonly="readonly">
&lt;%If Session("UserName")&lt;>"" Then%>
&lt;table width="100%" border="0" cellspacing="5" cellpadding="0">
		&lt;tr>
		  &lt;td colspan="2">���ã�&lt;%=Session("UserName")%>&lt;br />
	      &lt;a href="&lt;%=SiteDir%>User/Main.asp">�޸�����&lt;/a>&nbsp;|&nbsp;&lt;a href="&lt;%=SiteDir%>User/LoginOut.asp">�˳���¼&lt;/a>&lt;/td>
		  &lt;/tr>
&lt;/table>
&lt;%Else%>
&lt;table width="100%" border="0" cellspacing="5" cellpadding="0">
	&lt;form method="post" action="&lt;%=SiteDir%>User/Login.asp">
		&lt;tr>
		  &lt;td width="100" height="30" align="right">�û�����&lt;/td>
		  &lt;td>&lt;input name="UserName" type="text" class="ipt1" id="UserName" value="" maxlength="50" >&lt;/td>&lt;/tr>
		&lt;tr>
		  &lt;td height="30" align="right">���룺&lt;/td>
		  &lt;td>&lt;input name="PassWord" type="password" class="ipt1" id="PassWord" value="" maxlength="50">&lt;/td>
		&lt;/tr>
		&lt;tr>
		  &lt;td height="30" align="right">&lt;/td>
		  &lt;td>&lt;input name="Submit" type="submit" class="btn1" value="��¼" />&lt;/td>
	  &lt;/tr>
		&lt;tr>
		  &lt;td height="30" align="right">&lt;/td>
		  &lt;td>&lt;a href="&lt;%=SiteDir%>User/Register.asp">����ע��&lt;/a>  &lt;a href="&lt;%=SiteDir%>User/GetPass.asp">�һ�����&lt;/a>&lt;/td>
	  &lt;/tr>
	&lt;/form>
&lt;/table>
&lt;%End If%>
</textarea>
      <br>
      <input name="Submit" type="button" class="btn3" onClick="Copy('Code2')" value="�������ϴ���">
    �����ϴ���ŵ���Ҫ��ʾ��¼���λ��</td>
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