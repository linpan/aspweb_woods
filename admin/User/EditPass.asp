<!--#include file="../../Include/Class_Conn.asp"-->
<!--#include file="../../Include/Class_Main.asp"-->
<!--#include file="../../Include/Class_MD5.asp"-->
<!--#include file="../p8_Check.asp"-->
<%
	If Request.Form("PassWord")<>"" Then
		Dim PassWord,NewPassWord,RePassWord
		PassWord    = Request("PassWord")
		NewPassWord = Request("NewPassWord")
		RePassWord  = Request("RePassWord")

		If PassWord = "" Then
			Response.Write "<script>alert(""请填写当前密码"");window.history.back();</script>"
			Response.End()
		End If
		
		If NewPassWord = "" Then
			Response.Write "<script>alert(""请填写新密码"");window.history.back();</script>"
			Response.End()
		End If
		
		If Len(NewPassWord)<7 Then
			Response.Write "<script>alert(""新密码长度必须大于6位"");window.history.back();</script>"
			Response.End()
		End If
		
		If RePassWord = "" Then
			Response.Write "<script>alert(""请填写确认新密码"");window.history.back();</script>"
			Response.End()
		End If
		
		If NewPassWord <> RePassWord Then
			Response.Write "<script>alert(""确认新密码必须与新密码一致"");window.history.back();</script>"
			Response.End()
		End If
		
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.open "Select s_Pass From i6_Super Where s_User='"& Request.Cookies("Admin")("s_User") &"'",conn,1,3
		
		If Not rs.Eof Then
			If rs("s_Pass") <> Md5(PassWord) Then
				Response.Write "<script>alert(""当前密码错误"");window.history.back();</script>"
				Response.End()
			Else
				rs("s_Pass") = Md5(NewPassWord)
				rs.Update
			End If
		Else
			Response.Write "<script>alert(""用户不存在"");window.history.back();</script>"
			Response.End()
		End If

		CloseRs	
		CloseConn		
		
		Response.Write "<script>alert(""修改成功！"");window.close();</script>"
		Response.End()
		
	End If
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta http-equiv="pragma" content="no-cache" />
<meta http-equiv="Cache-Control" content="no-cache, must-revalidate" />
<meta http-equiv="expires" content="Wed, 26 Feb 1997 08:21:57 GMT" />
<meta http-equiv="expires" content="0" />
<title>修改密码</title>
<link href="../css/Public.css" rel="stylesheet" type="text/css" />
<script type="text/javascript">
<!--
function Trim(strSource){return strSource.replace(/^\s*/,'').replace(/\s*$/,'');}
function CheckForm(){
	if (Trim(document.getElementById("PassWord").value) == "") {
		alert("请填写 [当前密码]");
		document.getElementById("PassWord").focus();
		return false;
	}
	if (Trim(document.getElementById("NewPassWord").value) == "") {
		alert("请填写 [新密码]");
		document.getElementById("NewPassWord").focus();
		return false;
	}
	if (document.getElementById("NewPassWord").value.length <6) {
		alert("[新密码] 长度必须大于6");
		document.getElementById("NewPassWord").focus();
		return false;
	}
	if (document.getElementById("pwdMed").className=="" && document.getElementById("pwdHi").className=="") {
		alert("[新密码] 安全等级太低，请试用数字、字母、特殊字符组合的密码");
		document.getElementById("NewPassWord").focus();
		return false;
	}
	if (Trim(document.getElementById("RePassWord").value) == "") {
		alert("请填写 [确认新密码]");
		document.getElementById("RePassWord").focus();
		return false;
	}
	if (Trim(document.getElementById("RePassWord").value) != Trim(document.getElementById("NewPassWord").value)) {
		alert("[新密码] 必须与 [确认新密码] 一致");
		document.getElementById("RePassWord").focus();
		return false;
	}
}

function checkPassword(pwd){
	var objLow=document.getElementById("pwdLow");
	var objMed=document.getElementById("pwdMed");
	var objHi=document.getElementById("pwdHi");
	if(pwd.length>0){
	var p1= (pwd.search(/[a-zA-Z]/)!=-1) ? 1 : 0;
	var p2= (pwd.search(/[0-9]/)!=-1) ? 1 : 0;
	var p3= (pwd.search(/[^A-Za-z0-9_]/)!=-1) ? 1 : 0;
	var pa=p1+p2+p3;
	if(pa==1){
		objLow.className="pwdLight";
		objMed.className="";
		objHi.className="";
	}else if(pa==2){
		objLow.className="";
		objMed.className="pwdLight";
		objHi.className="";
	}else if(pa==3){
		objLow.className="";
		objMed.className="";
		objHi.className="pwdLight";
	}
	}
}
-->
</script>
<style type="text/css">
    .pwdLight {background-color:#FF3300; color:#fff;}
</style>
</head>
<body>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="30" bgcolor="#eaf3fd" style="border-bottom:1px solid #b5cef0;">&nbsp;&nbsp;<strong>修改密码</strong></td>
  </tr>
</table>
<form name="form1" method="post" action="EditPass.asp" onSubmit="return CheckForm();">
<table width="100%" height="100%" border="0" align="left" cellpadding="5" cellspacing="1">
      <tr bgcolor="#E4EDF9">
        <td height="30" align="right" bgcolor="#F8FBFE" width="100">当前密码：</td>
        <td colspan="9" bgcolor="#F8FBFE"><input name="PassWord" type="password" class="uIpt1" id="PassWord" size="20" maxlength="20" /></td>
      </tr>
      <tr bgcolor="#E4EDF9">
        <td height="30" align="right" bgcolor="#F8FBFE">新密码：</td>
        <td colspan="9" bgcolor="#F8FBFE" class="cBlack" style="line-height:150%;"><input name="NewPassWord" type="password" class="uIpt1" id="NewPassWord" size="20" maxlength="20" onKeyUp="checkPassword(this.value);" /></td>
      </tr>
      <tr bgcolor="#E4EDF9">
        <td height="30" align="right" bgcolor="#F8FBFE">密码安全等级：</td>
        <td colspan="9" bgcolor="#F8FBFE"><table width="200" border="0" cellspacing="0" cellpadding="0" style="border:1px solid #ddd; color:#a7a7a7;">
            <tr>
              <td align="center" id="pwdLow">低</td>
              <td align="center" id="pwdMed">中</td>
              <td align="center" id="pwdHi">高</td>
            </tr>
          </table></td>
      </tr>
      <tr bgcolor="#E4EDF9">
        <td height="30" align="right" bgcolor="#F8FBFE">重复新密码：</td>
        <td colspan="9" bgcolor="#F8FBFE"><input name="RePassWord" type="password" class="uIpt1" id="RePassWord" size="20" maxlength="20" /></td>
      </tr>
      <tr bgcolor="#E4EDF9">
        <td height="30" align="right" bgcolor="#F8FBFE">&nbsp;</td>
        <td colspan="9" bgcolor="#F8FBFE"><input type="submit" name="button" id="button" value="  提 交  "></td>
      </tr>
        <tr>
        <td colspan="10" bgcolor="#F8FBFE" style="padding-left:10px;">&nbsp;</td>
      </tr>
    </table>
</form>
</body>
</html>
<%
	CloseConn
%>