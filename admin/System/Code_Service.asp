<!--#include file="../../Include/Class_Conn.asp"-->
<!--#include file="../../Include/Class_Main.asp"-->
<% Super=1 'ֻ����������Ա���ʸ�ҳ %>
<!--#include file="../p8_Check.asp"-->
<%
	If Request("Submit") <> "" Then
		ServiceSwitch = Trim(Request("ServiceSwitch"))
		ServiceID     = Trim(Request("ServiceID"))
		ServiceCode   = Trim(Request("ServiceCode"))
		
		If ServiceSwitch="" Then
			Response.Write "<script>alert(""������ʧ���뷵��ˢ�£�"");window.history.back();</script>"
			Response.End()
		End If
		
		Set rs=Server.CreateObject("Adodb.Recordset")
		rs.Open "Select Top 1 ServiceSwitch,ServiceID,ServiceCode From p8_Config Order By id Asc",Conn,1,3
		
		If rs.Eof Then
			rs.AddNew
		End If
		
		rs("ServiceSwitch") = ServiceSwitch
		rs("ServiceID")     = ServiceID
		rs("ServiceCode")   = ServiceCode

		rs.Update
		rs.Close
		Set rs=Nothing
		
		CloseConn
		Response.Redirect "Code_Service.asp?Tip=���óɹ���"
		Response.End()
		
	End If
	
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>���߿ͷ� - ���ݵ���</title>
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
<%
	Dim Tip
	Tip = Request.QueryString("Tip")
	If Tip <> "" Then
		Response.Write "<script type=""text/javascript"">window.onload=function(){new x.creat(1, 41, 5, 10, '"& Tip &"');}</script>"
	End If
%>
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
    <tr>
      <td bgcolor="#F8FBFE"><table border="0" cellspacing="10" cellpadding="0">
        <tr>
          <td width="80" height="30" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Code_News_List.asp';">�����б�</td>
          <td width="80" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Code_News_View.asp';">������ʾ</td>
          <td width="80" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Code_Pic_List.asp';">ͼƬ�б�</td>
          <td width="80" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Code_Pic_View.asp';">ͼƬ��ʾ</td>
          <td width="80" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Code_Page.asp';">��ҳ</td>
          <td width="80" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Code_Form.asp';">��</td>
          <td width="80" align="center" class="Tab1_over">���߿ͷ�</td>
          <td width="80" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Code_User.asp';">��¼��</td>
        </tr>
      </table>	  </td>
    </tr>
    <tr>
      <td bgcolor="#F8FBFE" style="padding:10px;"><table width="100%" border="0" cellpadding="10" cellspacing="1" bgcolor="#E4EDF9">
          <tr>
            <td bgcolor="#FFFFFF" style="line-height:160%;"><strong>����˵����<br>
            </strong><span class="cGray">����Ҫ��ʾ���߿ͷ���ҳ����ô��롣</span></td>
          </tr>
      </table></td>
    </tr>
  </table>

<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
<%
	Dim ServiceSwitch,ServiceID
	Set rs=Server.CreateObject("Adodb.Recordset")
	rs.Open "Select Top 1 ServiceSwitch,ServiceID,ServiceCode From p8_Config Order By id Asc",Conn,1,3
	
	If Not rs.Eof Then
		ServiceSwitch = rs("ServiceSwitch")
		ServiceID     = rs("ServiceID")
		ServiceCode   = rs("ServiceCode")
	End If

	rs.Close
	Set rs=Nothing
%>
<form method="post" action="">
  <tr>
    <td height="30" align="right" bgcolor="#F8FBFE">״̬��&nbsp;</td>
    <td bgcolor="#F8FBFE" class="cBlack"><input name="ServiceSwitch" type="radio" value="1" <%If ServiceSwitch="1" Then Response.Write " checked=""checked"""%>>��&nbsp;&nbsp;
	<input name="ServiceSwitch" type="radio" value="0" <%If ServiceSwitch="0" Then Response.Write " checked=""checked"""%>>��</td>
  </tr>
  <tr>
    <td height="30" align="right" bgcolor="#F8FBFE">�ͷ��ʺţ�&nbsp;</td>
    <td bgcolor="#F8FBFE"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="500" valign="top"><textarea name="ServiceID" id="ServiceID" class="ipt3" style="width:500px; height:150px;"><%=ServiceID%></textarea></td>
          <td valign="top" style="padding-left:10px; line-height:150%; color:#999;">��д���¼�����������ʺţ�һ��һ������ʽ���£�<BR>
            QQ = 136310631<br>
            MSN = psd8@hotmail.com<br>
            ���� = psd8<br>
			skype = psd8<br>
			�ٶ�hi = psd8<br>			</td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td height="30" align="right" bgcolor="#F8FBFE">�������룺&nbsp;</td>
    <td bgcolor="#F8FBFE"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="500" valign="top"><textarea name="ServiceCode" id="ServiceCode" class="ipt3" style="width:500px; height:80px;"><%=ServiceCode%></textarea></td>
          <td valign="top" style="padding-left:10px; line-height:150%; color:#999;">�˴��ɷ��õ������ͷ����룬��ͨ�����´�����е��ã�����ͳһ����</td>
        </tr>
    </table></td>
  </tr>

  <tr>
  <td width="74" height="30" align="center" bgcolor="#F8FBFE">&nbsp;</td>
  <td bgcolor="#F8FBFE"><span style="padding:20px 0;">
    <input name="Submit" type="submit" class="btn2" value="��������">
  </span></td>
</tr>
<tr>
  <td height="30" align="center" bgcolor="#F8FBFE">&nbsp;</td>
  <td bgcolor="#F8FBFE">&nbsp;</td>
</tr>
<tr>
  <td height="30" align="center" bgcolor="#F8FBFE">���ô��룺</td>
  <td bgcolor="#F8FBFE"><textarea id="Code1" class="ipt3" style="width:500px; height:80px;" readonly="readonly"><script type="text/javascript" src="<%=SiteDir%>Include/Service/Class_Safe.asp"></script></textarea>
      <br>
      <input name="Submit2" type="button" class="btn3" onClick="Copy('Code1')" value="�������ϴ���">
    �����ϴ���ŵ�&lt;body&gt;&lt;/body&gt;��</td>
</tr>
<tr>
  <td height="30" align="center" bgcolor="#F8FBFE">&nbsp;</td>
  <td bgcolor="#F8FBFE">&nbsp;</td>
</tr>
</form>
</table>

</body>
</html>
<%
	CloseConn
%>