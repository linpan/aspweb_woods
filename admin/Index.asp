<!--#include file="../Include/Class_Conn.asp"-->
<!--#include file="../Include/Class_Main.asp"-->
<!--#include file="../Include/Class_MD5.asp"-->
<%
	Dim s_User,s_Pass,rs
	s_User = Replace_Text(Request.Form("p8_User"))
	s_Pass = Request.Form("p8_Pass")
	
	If s_User<>"" And s_Pass<>"" Then

		Check_url()	
		Set rs = Server.CreateObject ("ADODB.Recordset")
		rs.Open "Select s_User,s_Pass,s_Name,s_IP,s_Date,s_Error,s_ErrorDate,s_Level From p8_Super Where s_User='"& s_User &"'",Conn,1,1

		If Not rs.Eof Then
			
			If DateDiff("n",rs("s_ErrorDate"),Now)>SysLgLock Then Conn.Execute = "Update p8_Super Set s_Error=0 Where s_User='"& s_User &"'" '�����½�������10���ӣ�����������0
			
			If rs("s_Error")>=SysLgErr And DateDiff("n",rs("s_ErrorDate"),Now)<=SysLgLock Then '�����½����5�����ϣ���10�����ڲ������½
				Response.Write "<script>alert(""�ʺ��ѱ����ã��޷���½��"");window.history.back();</script>"
				Response.End()
			Else
				If rs("s_Pass") <> MD5(s_Pass) Then
					conn.Execute = "Update p8_Super Set s_Error=s_Error+1,s_ErrorDate='"& Now() &"' Where s_User='"& s_User &"'" '�ۼƴ������
					Conn.Execute = "INSERT INTO p8_Log (UserName, LoginIP, LoginDate, UserLevel, LoginState) VALUES ('"& s_User &"', '"& Request.Servervariables("Remote_Addr") &"','"& Now() &"','"& rs("s_Level") &"','��¼ʧ��')" 'д���¼��־
					Response.Write "<script>alert(""�û������������"");window.history.back();</script>"
					Response.End()
				Else					
					'Response.Cookies("Admin").Domain = ""
					Response.Cookies("Admin").Expires= DateAdd("d",1,date) 'ָ��1������
					
					Response.Cookies("Admin")("s_User")  = s_User
					Response.Cookies("Admin")("s_Pass")  = MD5(s_Pass)
					Response.Cookies("Admin")("s_Name")  = rs("s_Name")
					Response.Cookies("Admin")("s_Level") = rs("s_Level")
					
'					Response.Cookies("Admin")("s_IP") = rs("s_IP")
'					Response.Cookies("Admin")("s_Date") = rs("s_Date")
					
					Conn.Execute = "Update p8_Super Set s_Date='"& Now() &"',s_IP='"& Request.Servervariables("Remote_Addr") &"',s_Count=s_Count+1,s_Error=0 Where s_User='"& s_User &"'" '�����û�����¼��Ϣ
					Conn.Execute = "INSERT INTO p8_Log (UserName, LoginIP, LoginDate, UserLevel, LoginState) VALUES ('"& s_User &"', '"& Request.Servervariables("Remote_Addr") &"','"& Now() &"','"& rs("s_Level") &"','��¼�ɹ�')" 'д���¼��־
					Conn.Execute "Delete From p8_Log Where datediff('d',LoginDate,Now())>30" 'ɾ��30��ǰ��־
					
					Response.Redirect "p8_Main.asp"
					Response.End()
				End If
			End If
		Else
			Response.Write "<script>alert(""�û������������"");window.history.back();</script>"
			Response.End()
		End If
		
		CloseRs
		CloseConn
	End If
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��̨��������</title>
<meta name="robots" content="noindex,nofollow" />
<link href="css/Public.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
	body {background:url(images/Login/bg.jpg) center 0 #7fcaf5 no-repeat;}
	input {border:none; background:none; font-size:13px; color:#ff5400;}
-->
</style>
<script type="text/javascript">
<!--
function Trim(strSource){return strSource.replace(/^\s*/,'').replace(/\s*$/,'');}
function CheckForm(){
	if (Trim(document.getElementById("p8_User").value) == "") {
		alert("�������ʺţ�");
		document.getElementById("p8_User").focus();
		return false;
	}
	if (Trim(document.getElementById("p8_Pass").value) == "") {
		alert("���������룡");
		document.getElementById("p8_Pass").focus();
		return false;
	}
}
-->
</script>
</head>
<body scroll="no">
<form name="mbLogin" method="post" action="Index.asp" onSubmit="return CheckForm();">
  <table width="200" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td height="266" colspan="3">&nbsp;</td>
    </tr>
    <tr>
      <td width="50" height="50">&nbsp;</td>
      <td colspan="2" align="center"><input name="p8_User" type="text" id="p8_User" style="width:145px;" value="admin" maxlength="50"></td>
    </tr>
    <tr>
      <td height="8" colspan="3"></td>
    </tr>
    <tr>
      <td height="29">&nbsp;</td>
      <td colspan="2" align="center"><input name="p8_Pass" type="password" id="p8_Pass" style="width:145px;" value="admin" maxlength="50"></td>
    </tr>
    <tr>
      <td height="9" colspan="3"></td>
    </tr>
    <tr>
      <td height="53" colspan="3" align="center"><input type="submit" name="Submit" value="" style="width:122px; height:37px; background:url(images/Login/Lg_Btn.jpg) no-repeat; cursor:pointer;"></td>
    </tr>
  </table>
  <script>//mbLogin.submit();</script>
</form>
<span style="display:none;"></span>
</body>
</html>