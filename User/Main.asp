<!--#include file="../Include/Class_Conn.asp"-->
<!--#include file="../Include/Class_Main.asp"-->
<!--#include file="../Include/Class_MD5.asp"-->
<!--#include file="LoginCheck.asp"-->
<%
	If Request("Submit")<>"" Then
		PassWord   = Replace_Text(Trim(Request.Form("PassWord")))
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

		If Email = "" Then
			Response.Write "<script>alert(""����дEmail"");window.history.back();</script>"
			Response.End()
		End If
		
		Set rs = Server.Createobject("Adodb.RecordSet")
		rs.open "Select id From p8_User Where Email ='"& Email &"' And UserName<>'"& Session("UserName") &"'",Conn,1,1
		If Not rs.Eof Then
			Response.Write "<script>alert(""Email�ѱ�ע�ᣬ��ʹ������Email"");window.history.back();</script>"
			Response.End()
		End If
		rs.Close
		Set rs = Nothing

		Set rs=Server.CreateObject("Adodb.Recordset")
		rs.Open "Select * From p8_User Where UserName = '"& Session("UserName") &"'",Conn,1,3

		If PassWord<>"" Then rs("PassWord") = MD5(PassWord)
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

		rs.Update
		rs.Close
		Set rs=Nothing
		Response.Write "<script>alert(""�޸ĳɹ�"& Msg &"��"");window.location.href=window.location.href;</script>"
	End If
%>
<script type="text/javascript">
function $p8(Obj){return document.getElementById(Obj);}
function Trim(strSource){return strSource.replace(/^\s*/,'').replace(/\s*$p8/,'');}
function CheckForm(){
	if(Trim($p8("Email").value) == ''){
		alert("����дEmail");
		$p8("Email").focus();
		return false;
	}
	if(Trim($p8("RePassWord").value) != Trim($p8("PassWord").value)){
		alert("�ظ����������������ͬ");
		$p8("RePassWord").focus();
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
<title>��Ա����</title>
<link href="css/Public.css" rel="stylesheet" type="text/css" />
</head>

<body>
<%
	Set rs = Server.CreateObject ("ADODB.Recordset")
	rs.Open "Select * From p8_User Where UserName ='"& Session("UserName") &"'",Conn,1,1
%>
<div style="width:560px; margin-top:50px; padding:20px 20px 20px 20px; background:url(images/User_Top.jpg) no-repeat; border:1px solid #1b72af;">
	<table width="100%" border="0" cellspacing="5" cellpadding="0">
		<form method="post" action="" onSubmit="return CheckForm()">
			<tr>
			  <td height="30" align="right">&nbsp;</td>
			  <td align="right"><a href="LoginOut.asp">�˳���¼</a> | <a href="<%=SiteDir%>">������ҳ&gt;&gt;</a></td>
		  </tr>
			<tr>
			  <td height="30" colspan="2" style="border:1px solid #F2E4AE; background-color:#FDFEE9; line-height:180%; padding:10px; color:#333;">���ã�<font color="#FF0000"><%=rs("UserName")%></font><br />
			    <strong>�ϴε�¼ʱ�䣺</strong><%=rs("LoginDate")%><br />
			    <strong>�ϴε�¼IP��</strong><%=rs("LoginIP")%><br />
			</td>
		  </tr>
			<tr>
			  <td width="100" height="30" align="right">�û�����</td>
			  <td><%=rs("UserName")%></td></tr>
			<tr>
			  <td height="30" align="right">���룺</td>
			  <td><input name="PassWord" type="password" class="ipt1" id="PassWord" value="" maxlength="50">
		      &nbsp;<font color="#999999">�粻�޸ģ��뱣��Ϊ��</font></td>
			</tr>
			<tr>
			  <td height="30" align="right">�ظ����룺</td>
			  <td><input name="RePassWord" type="password" class="ipt1" id="RePassWord" value="" maxlength="50"></td>
			</tr>
			<tr>
			  <td height="30" align="right">Email��</td>
			  <td><input name="Email" type="text" class="ipt1" id="Email" value="<%=rs("Email")%>" maxlength="50" /></td>
			</tr>
			<tr>
			  <td height="30" align="right">������  </td>
			  <td><input name="FullName" type="text" class="ipt1" id="FullName" value="<%=rs("FullName")%>" maxlength="10" /></td>
			</tr>
			<tr>
			  <td height="30" align="right">�Ա�  </td>
			  <td><input type="radio" name="Sex" value="��" <%If rs("Sex") = "��" Then Response.Write " checked=""checked"""%> />��
			  <input type="radio" name="Sex" value="Ů" <%If rs("Sex") = "Ů" Then Response.Write " checked=""checked"""%> />Ů</td>
			</tr>
			<tr>
			  <td height="30" align="right">��˾���ƣ�  </td>
			  <td><input name="Company" type="text" class="ipt1" id="Company" value="<%=rs("Company")%>" maxlength="100"></td>
			</tr>
			<tr>
			  <td height="30" align="right">�绰��  </td>
			  <td><input name="Tel" type="text" class="ipt1" id="Tel" value="<%=rs("Tel")%>" maxlength="50"></td>
			</tr>
			<tr>
			  <td height="30" align="right">���棺  </td>
			  <td><input name="Fax" type="text" class="ipt1" id="Fax" value="<%=rs("Fax")%>" maxlength="50"></td>
			</tr>
			<tr>
			  <td height="30" align="right">�ֻ���  </td>
			  <td><input name="Mob" type="text" class="ipt1" id="Mob" value="<%=rs("Mob")%>" maxlength="50"></td>
			</tr>
			<tr>
			  <td height="30" align="right">��ϵ��ַ��  </td>
			  <td><input name="Address" type="text" class="ipt1" id="Address" value="<%=rs("Address")%>" maxlength="100"></td>
			</tr>
			<tr>
			  <td height="30" align="right">�ʱࣺ  </td>
			  <td><input name="Zipcode" type="text" class="ipt1" id="Zipcode" value="<%=rs("Zipcode")%>" maxlength="10"></td>
			</tr>
			<tr>
			  <td height="30" align="right">��ַ��  </td>
			  <td><input name="Url" type="text" class="ipt1" id="Url" value="<%=rs("Url")%>" maxlength="50"></td>
			</tr>
			<tr>
			  <td width="100" height="30" align="right"> </td>
			  <td><input name="Submit" type="submit" class="btn1" value="�޸�" /></td>
			</tr>
		</form>
	</table>
</div>
<%
	CloseRs
	CloseConn
%>
</body>
</html>
