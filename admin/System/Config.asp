<!--#include file="../../Include/Class_Conn.asp"-->
<!--#include file="../../Include/Class_Main.asp"-->
<% Super=1 'ֻ����������Ա���ʸ�ҳ %>
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
			Response.Write "<script>alert(""����д��վ����Ŀ¼"");window.history.back();</script>"
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
		Response.Redirect "Config.asp?Tip=���óɹ���"
		Response.End()
		
	End If
	
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��վ����</title>
<script type="text/javascript">top.window.aTitle.innerText='��վ����'</script>
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
		alert("����д��վ����Ŀ¼");
		$p8("SiteDir").focus();
		return false;
	}
	if(Trim($p8("FsoName").value) == ''){
		alert("����д������FSO�����");
		$p8("FsoName").focus();
		return false;
	}
	if(Trim($p8("UserLgErr").value) == ''){
		alert("����д��Ա��¼����������");
		$p8("UserLgErr").focus();
		return false;
	}
	if(Trim($p8("UserLgLock").value) == ''){
		alert("����д�ﵽ�����������ʱ��");
		$p8("UserLgLock").focus();
		return false;
	}
	if(Trim($p8("SysLgErr").value) == ''){
		alert("����д��̨��¼����������");
		$p8("SysLgErr").focus();
		return false;
	}
	if(Trim($p8("SysLgLock").value) == ''){
		alert("����д�ﵽ�����������ʱ��");
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
		Response.Write "���ݿ�ϵͳ������ʧ���������ݿ⣡"
		Response.End()
	End If
%>
<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
<form method="post" action="Config.asp" onSubmit="return CheckForm()">
  <tr>
    <td width="150" height="30" align="right" bgcolor="#F8FBFE">��վ����Ŀ¼��</td>
    <td bgcolor="#F8FBFE"><input name="SiteDir" type="text" class="ipt3" id="SiteDir" value="<%=rs("SiteDir")%>" maxlength="50" style="width:200px;">
      <span class="cGray">��ʼ�ͽ�β�����/</span></td>
  </tr>
  <tr>
    <td height="30" align="right" bgcolor="#F8FBFE">������FSO�������</td>
    <td bgcolor="#F8FBFE"><input name="FsoName" type="text" class="ipt3" id="FsoName" value="<%=rs("FsoName")%>" maxlength="50" style="width:200px;"></td>
  </tr>
  <tr>
    <td height="30" align="right" bgcolor="#F8FBFE">��Աע����ˣ�</td>
    <td bgcolor="#F8FBFE"><input name="UserCheck" type="checkbox" id="UserCheck" value="1" <%If rs("UserCheck")=1 Then Response.Write "checked=""checked"""%>>��Ҫ���</td>
  </tr>
  <tr>
    <td height="30" align="right" bgcolor="#F8FBFE">��Ա��¼������������</td>
    <td bgcolor="#F8FBFE"><input name="UserLgErr" type="text" class="ipt3" id="UserLgErr" value="<%=rs("UserLgErr")%>" maxlength="6" style="width:200px;"></td>
  </tr>
  <tr>
    <td height="30" align="right" bgcolor="#F8FBFE">�ﵽ�����������ʱ����</td>
    <td bgcolor="#F8FBFE"><input name="UserLgLock" type="text" class="ipt3" id="UserLgLock" value="<%=rs("UserLgLock")%>" maxlength="6" style="width:200px;">
      ���� <span class="cGray">��ע���Ա��¼�����ﵽ������ʾ������ٷ����ڲ������¼</span> </td>
  </tr>
  <tr>
    <td height="30" align="right" bgcolor="#F8FBFE">��̨��¼������������</td>
    <td bgcolor="#F8FBFE"><input name="SysLgErr" type="text" class="ipt3" id="SysLgErr" value="<%=rs("SysLgErr")%>" maxlength="6" style="width:200px;"></td>
  </tr>
  <tr>
    <td height="30" align="right" bgcolor="#F8FBFE">�ﵽ�����������ʱ����</td>
    <td bgcolor="#F8FBFE"><input name="SysLgLock" type="text" class="ipt3" id="SysLgLock" value="<%=rs("SysLgLock")%>" maxlength="6" style="width:200px;">
      ���� <span class="cGray">����̨�û���¼�����ﵽ������ʾ������ٷ����ڲ������¼</span> </td>
  </tr>
  <tr>
    <td height="30" align="right" bgcolor="#F8FBFE">SMTP�ʼ���ַ��</td>
    <td bgcolor="#F8FBFE"><input name="SmtpEmail" type="text" class="ipt3" id="SmtpEmail" value="<%=rs("SmtpEmail")%>" maxlength="50" style="width:200px;">
      &nbsp;<span class="cGray">�������û������ʼ�����ʽ�磺moumou@qq.com</span></td>
  </tr>
  <tr>
    <td height="30" align="right" bgcolor="#F8FBFE">SMTP�ʼ��û�����</td>
    <td bgcolor="#F8FBFE"><input name="SmtpUser" type="text" class="ipt3" id="SmtpUser" value="<%=rs("SmtpUser")%>" maxlength="50" style="width:200px;">
      &nbsp;<span class="cGray">��ʽ�磺moumou</span></td>
  </tr>
  <tr>
    <td height="30" align="right" bgcolor="#F8FBFE">SMTP�ʼ����룺</td>
    <td bgcolor="#F8FBFE"><input name="SmtpPass" type="text" class="ipt3" id="SmtpPass" value="<%=rs("SmtpPass")%>" maxlength="50" style="width:200px;"></td>
  </tr>
  <tr>
    <td height="30" align="right" bgcolor="#F8FBFE">SMTP��������</td>
    <td bgcolor="#F8FBFE"><input name="SmtpServer" type="text" class="ipt3" id="SmtpServer" value="<%=rs("SmtpServer")%>" maxlength="50" style="width:200px;">
      &nbsp;<span class="cGray">��ʽ�磺smtp.qq.com</span></td>
  </tr>
  <tr>
    <td height="30" align="right" bgcolor="#F8FBFE">&nbsp;</td>
    <td bgcolor="#F8FBFE"><span style="padding:20px 0;">
      <input name="Submit" type="submit" class="btn2" value=" ���� " >
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