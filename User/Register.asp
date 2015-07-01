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
			Response.Write "<script>alert(""请填写用户名"");window.history.back();</script>"
			Response.End()
		End If
		
		If PassWord = "" Then
			Response.Write "<script>alert(""请填写密码"");window.history.back();</script>"
			Response.End()
		End If
		
		If Email = "" Then
			Response.Write "<script>alert(""请填写Email"");window.history.back();</script>"
			Response.End()
		End If
		
		Set rs = Server.Createobject("Adodb.RecordSet")
		rs.open "Select id From p8_User Where UserName ='"& UserName &"'",Conn,1,1
		If Not rs.Eof Then
			Response.Write "<script>alert(""用户名已被注册，请使用其他用户名"");window.history.back();</script>"
			Response.End()
		End If
		rs.Close
		Set rs = Nothing
		
		Set rs = Server.Createobject("Adodb.RecordSet")
		rs.open "Select id From p8_User Where Email ='"& Email &"'",Conn,1,1
		If Not rs.Eof Then
			Response.Write "<script>alert(""Email已被注册，请使用其他Email"");window.history.back();</script>"
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
			Msg = ",请等待管理员审核"
		End If

		rs.Update
		rs.Close
		Set rs=Nothing
		Response.Write "<script>alert(""注册成功"& Msg &"！"");window.location.href='Main.asp';</script>"
	End If
%>
<script type="text/javascript">
function $p8(Obj){return document.getElementById(Obj);}
function Trim(strSource){return strSource.replace(/^\s*/,'').replace(/\s*$p8/,'');}
function CheckForm(){
	if(Trim($p8("UserName").value) == ''){
		alert("请填写用户名");
		$p8("UserName").focus();
		return false;
	}
	if(Trim($p8("UserName").value).length < 3){
		alert("用户名长度不能小于3位");
		$p8("UserName").focus();
		return false;
	}
	var re=/^[\da-zA-Z_]+$/;
	if(!re.test($p8("UserName").value)){
		alert("用户名应由字母、数字和下划线组成");
		$p8("UserName").focus();
		return false;
	}
	if(Trim($p8("PassWord").value) == ''){
		alert("请填写密码");
		$p8("PassWord").focus();
		return false;
	}
	if(Trim($p8("RePassWord").value) != Trim($p8("PassWord").value)){
		alert("重复密码必须与密码相同");
		$p8("RePassWord").focus();
		return false;
	}
	if(Trim($p8("Email").value) == ''){
		alert("请填写Email");
		$p8("Email").focus();
		return false;
	}
	if(!CheckEmail($p8("Email").value)){
		alert("Email格式错误");
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
<title>会员注册</title>
<link href="css/Public.css" rel="stylesheet" type="text/css" />
</head>

<body>
<div style="width:560px; margin-top:50px; padding:60px 20px 20px 20px; background:url(images/Reg_Top.jpg) no-repeat; border:1px solid #1b72af;">
	<table width="100%" border="0" cellspacing="5" cellpadding="0">
		<form method="post" action="" onSubmit="return CheckForm()">
			<tr>
			  <td width="100" height="30" align="right">用户名：<font color="#FF3300">*</font> </td>
			  <td><input name="UserName" type="text" class="ipt1" id="UserName" value="" maxlength="50" ></td></tr>
			<tr>
			  <td height="30" align="right">密码：<font color="#FF3300">*</font> </td>
			  <td><input name="PassWord" type="password" class="ipt1" id="PassWord" value="" maxlength="50"></td>
			</tr>
			<tr>
			  <td height="30" align="right">重复密码：<font color="#FF3300">*</font> </td>
			  <td><input name="RePassWord" type="password" class="ipt1" id="RePassWord" value="" maxlength="50"></td>
			</tr>
			<tr>
			  <td height="30" align="right">Email：<font color="#FF3300">*</font> </td>
			  <td><input name="Email" type="text" class="ipt1" id="Email" value="" maxlength="50" /></td>
			</tr>
			<tr>
			  <td height="30" align="right">姓名：  </td>
			  <td><input name="FullName" type="text" class="ipt1" id="FullName" value="" maxlength="10" /></td>
			</tr>
			<tr>
			  <td height="30" align="right">性别：  </td>
			  <td><input type="radio" name="Sex" value="男" />男
			  <input type="radio" name="Sex" value="女" />女</td>
			</tr>
			<tr>
			  <td height="30" align="right">公司名称：  </td>
			  <td><input name="Company" type="text" class="ipt1" id="Company" value="" maxlength="100"></td>
			</tr>
			<tr>
			  <td height="30" align="right">电话：  </td>
			  <td><input name="Tel" type="text" class="ipt1" id="Tel" value="" maxlength="50"></td>
			</tr>
			<tr>
			  <td height="30" align="right">传真：  </td>
			  <td><input name="Fax" type="text" class="ipt1" id="Fax" value="" maxlength="50"></td>
			</tr>
			<tr>
			  <td height="30" align="right">手机：  </td>
			  <td><input name="Mob" type="text" class="ipt1" id="Mob" value="" maxlength="50"></td>
			</tr>
			<tr>
			  <td height="30" align="right">联系地址：  </td>
			  <td><input name="Address" type="text" class="ipt1" id="Address" value="" maxlength="100"></td>
			</tr>
			<tr>
			  <td height="30" align="right">邮编：  </td>
			  <td><input name="Zipcode" type="text" class="ipt1" id="Zipcode" value="" maxlength="10"></td>
			</tr>
			<tr>
			  <td height="30" align="right">网址：  </td>
			  <td><input name="Url" type="text" class="ipt1" id="Url" value="" maxlength="50"></td>
			</tr>
			<tr>
			  <td width="100" height="30" align="right"> </td>
			  <td><input name="Submit" type="submit" class="btn1" value="注册" /></td>
			</tr>
		</form>
	</table>
</div>
<%
	CloseConn
%>
</body>
</html>
