<!--#include file="../Include/Class_Conn.asp"-->
<!--#include file="../Include/Class_Main.asp"-->
<!--#include file="../Include/Class_MD5.asp"-->
<%
	If Request("UserName")<>"" Then
		UserName   = Replace_Text(Trim(Request.Form("UserName")))
		PassWord   = Replace_Text(Trim(Request.Form("PassWord")))
		RePassWord = Replace_Text(Trim(Request.Form("RePassWord")))
		Email      = Replace_Text(Trim(Request.Form("Email")))
		FullName   = Replace_Text(Trim(Request.Form("FullName")))
		Sex        = Replace_Text(Trim(Request.Form("Sex")))
		Company    = Replace_Text(Trim(Request.Form("Company")))
		Tel        = Replace_Text(Trim(Request.Form("Tel")))
		Fax        = Replace_Text(Trim(Request.Form("Fax")))
		Mob        = Replace_Text(Trim(Request.Form("Mob")))
		Address    = Replace_Text(Trim(Request.Form("Address")))
		Zipcode    = Replace_Text(Trim(Request.Form("Zipcode")))
		Url        = Replace_Text(Trim(Request.Form("Url"))	)

		If UserName = "" Then
			Response.Write "<script>alert(""����д�û���"");window.history.back();</script>"
			Response.End()
		End If
		
		If PassWord = "" Then
			Response.Write "<script>alert(""����д����"");window.history.back();</script>"
			Response.End()
		End If
		
		If Email = "" Then
			Response.Write "<script>alert(""����дEmail"");window.history.back();</script>"
			Response.End()
		End If
		
		Set rs = Server.Createobject("Adodb.RecordSet")
		rs.open "Select id From p8_User Where UserName ='"& UserName &"'",Conn,1,1
		If Not rs.Eof Then
			Response.Write "<script>alert(""�û����ѱ�ע�ᣬ��ʹ�������û���"");window.history.back();</script>"
			Response.End()
		End If
		rs.Close
		Set rs = Nothing
		
		Set rs = Server.Createobject("Adodb.RecordSet")
		rs.open "Select id From p8_User Where Email ='"& Email &"'",Conn,1,1
		If Not rs.Eof Then
			Response.Write "<script>alert(""Email�ѱ�ע�ᣬ��ʹ������Email"");window.history.back();</script>"
			Response.End()
		End If
		rs.Close
		Set rs = Nothing

		Set rs=Server.CreateObject("Adodb.Recordset")
		rs.Open "Select * From p8_User",Conn,1,3
		rs.AddNew

		rs("UserName") = UserName
		rs("PassWord") = MD5(PassWord)
		rs("Email")    = Email
		rs("FullName") = FullName
		rs("Sex")      = Sex
		rs("Company")  = Company
		rs("Tel")      = Tel
		rs("Fax")      = Fax
		rs("Mob")      = Mob
		rs("Address")  = Address
		rs("Zipcode")  = Zipcode
		rs("Url")      = Url
		rs("LoginIP")  = Request.Servervariables("Remote_Addr")
		If UserCheck = 1 Then
			rs("UserState") = 0
			Msg = ",��ȴ�����Ա���"
		End If

		rs.Update
		rs.Close
		Set rs=Nothing
		Response.Write "<script>alert(""ע��ɹ�"& Msg &"��"");window.location.href='Main.asp';</script>"
	End If
%>
<script type="text/javascript">
function $p8(Obj){return document.getElementById(Obj);}
function Trim(strSource){return strSource.replace(/^\s*/,'').replace(/\s*$p8/,'');}
function CheckForm(){
	if(Trim($p8("UserName").value) == ''){
		alert("����д�û���");
		$p8("UserName").focus();
		return false;
	}
	if(Trim($p8("UserName").value).length < 3){
		alert("�û������Ȳ���С��3λ");
		$p8("UserName").focus();
		return false;
	}
	var re=/^[\da-zA-Z_]+$/;
	if(!re.test($p8("UserName").value)){
		alert("�û���Ӧ����ĸ�����ֺ��»������");
		$p8("UserName").focus();
		return false;
	}
	if(Trim($p8("PassWord").value) == ''){
		alert("����д����");
		$p8("PassWord").focus();
		return false;
	}
	if(Trim($p8("RePassWord").value) != Trim($p8("PassWord").value)){
		alert("�ظ����������������ͬ");
		$p8("RePassWord").focus();
		return false;
	}
	if(Trim($p8("Email").value) == ''){
		alert("����дEmail");
		$p8("Email").focus();
		return false;
	}
	if(!CheckEmail($p8("Email").value)){
		alert("Email��ʽ����");
		$p8("Email").focus();
		return false;
	}
	return true;
}
function CheckEmail(e){
	var ok = "1234567890qwertyuiop[]asdfghjklzxcvbnm.+@-_QWERTYUIOPASDFGHJKLZXCVBNM";
	for(var i=0; i<e.length; i++){
		if (ok.indexOf(e.charAt(i))<0) {
			return false;
		}
	}
	if(e.indexOf("@")<=0){
		return false;
	}
	if(e.indexOf(".")<=0){
		return false;
	}	
	return true;
}
</script>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��Աע��</title>
<link href="css/Public.css" rel="stylesheet" type="text/css" />
</head>

<body>
<div style="width:560px; margin-top:50px; padding:60px 20px 20px 20px; background:url(images/Reg_Top.jpg) no-repeat; border:1px solid #1b72af;">
	<table width="100%" border="0" cellspacing="5" cellpadding="0">
		<form method="post" action="" onSubmit="return CheckForm()">
			<tr>
			  <td width="100" height="30" align="right">�û�����<font color="#FF3300">*</font> </td>
			  <td><input name="UserName" type="text" class="ipt1" id="UserName" value="" maxlength="50" ></td></tr>
			<tr>
			  <td height="30" align="right">���룺<font color="#FF3300">*</font> </td>
			  <td><input name="PassWord" type="password" class="ipt1" id="PassWord" value="" maxlength="50"></td>
			</tr>
			<tr>
			  <td height="30" align="right">�ظ����룺<font color="#FF3300">*</font> </td>
			  <td><input name="RePassWord" type="password" class="ipt1" id="RePassWord" value="" maxlength="50"></td>
			</tr>
			<tr>
			  <td height="30" align="right">Email��<font color="#FF3300">*</font> </td>
			  <td><input name="Email" type="text" class="ipt1" id="Email" value="" maxlength="50" /></td>
			</tr>
			<tr>
			  <td height="30" align="right">������  </td>
			  <td><input name="FullName" type="text" class="ipt1" id="FullName" value="" maxlength="10" /></td>
			</tr>
			<tr>
			  <td height="30" align="right">�Ա�  </td>
			  <td><input type="radio" name="Sex" value="��" />��
			  <input type="radio" name="Sex" value="Ů" />Ů</td>
			</tr>
			<tr>
			  <td height="30" align="right">��˾���ƣ�  </td>
			  <td><input name="Company" type="text" class="ipt1" id="Company" value="" maxlength="100"></td>
			</tr>
			<tr>
			  <td height="30" align="right">�绰��  </td>
			  <td><input name="Tel" type="text" class="ipt1" id="Tel" value="" maxlength="50"></td>
			</tr>
			<tr>
			  <td height="30" align="right">���棺  </td>
			  <td><input name="Fax" type="text" class="ipt1" id="Fax" value="" maxlength="50"></td>
			</tr>
			<tr>
			  <td height="30" align="right">�ֻ���  </td>
			  <td><input name="Mob" type="text" class="ipt1" id="Mob" value="" maxlength="50"></td>
			</tr>
			<tr>
			  <td height="30" align="right">��ϵ��ַ��  </td>
			  <td><input name="Address" type="text" class="ipt1" id="Address" value="" maxlength="100"></td>
			</tr>
			<tr>
			  <td height="30" align="right">�ʱࣺ  </td>
			  <td><input name="Zipcode" type="text" class="ipt1" id="Zipcode" value="" maxlength="10"></td>
			</tr>
			<tr>
			  <td height="30" align="right">��ַ��  </td>
			  <td><input name="Url" type="text" class="ipt1" id="Url" value="" maxlength="50"></td>
			</tr>
			<tr>
			  <td width="100" height="30" align="right"> </td>
			  <td><input name="Submit" type="submit" class="btn1" value="ע��" /></td>
			</tr>
		</form>
	</table>
</div>
<%
	CloseConn
%>
</body>
</html>
