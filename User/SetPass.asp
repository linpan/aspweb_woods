<!--#include file="../Include/Class_Conn.asp"-->
<!--#include file="../Include/Class_Main.asp"-->
<!--#include file="../Include/Class_MD5.asp"-->
<%
	If Request.Form("PassWord")<>"" Then
		Check_url()
		Dim PassWord,RePassWord,PassKey,Tip
		PassKey     = Replace_Text(Request("PassKey"))
		PassWord    = Request("PassWord")
		RePassWord  = Request("RePassWord")

		If PassKey = "" Then
			Response.Write "<script>alert(""非法操作！"");window.history.back();</script>"
			Response.End()
		End If
		
		If PassWord = "" Then
			Response.Write "<script>alert(""请填写新密码"");window.history.back();</script>"
			Response.End()
		End If
		
		If Len(PassWord)<6 Then
			Response.Write "<script>alert(""新密码长度必须大于6位"");window.history.back();</script>"
			Response.End()
		End If
		
		If RePassWord = "" Then
			Response.Write "<script>alert(""请填写确认新密码"");window.history.back();</script>"
			Response.End()
		End If
		
		If PassWord <> RePassWord Then
			Response.Write "<script>alert(""确认新密码必须与新密码一致"");window.history.back();</script>"
			Response.End()
		End If
		
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.open "Select PassWord,PassKey From p8_User Where PassKey='"& PassKey &"' ",conn,1,3
		
		If Not rs.Eof Then
			rs("PassWord") = Md5(PassWord)
			rs("PassKey")  = ""
			rs.Update
		End If

		CloseRs	
		
		Tip = "<font color=""red"">恭喜！密码修改成功！</font><br /><a href="& SiteDir &">返回首页&gt;&gt;</a>"
		
	End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>找回密码</title>
<link href="css/Public.css" rel="stylesheet" type="text/css" />
</head>

<body>
<div style="width:560px; margin-top:50px; padding:60px 20px 20px 20px; background:url(images/Pass_Top.jpg) no-repeat; border:1px solid #1b72af;">
<%
	If Tip = "" Then
		Dim rs
		PassKey = Replace_Text(Request("K"))
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.open "Select PassKey From p8_User Where PassKey='"& PassKey &"'",conn,1,1
		
		If rs.Eof Or Trim(PassKey)="" Or isNull(PassKey) Then
%>
		<table width="100%" border="0" cellspacing="5" cellpadding="0">
          <tr>
            <td style="font-size:14px;">该地址已失效，如收到多封邮件，请使用最后一封邮件。<br /><br />或<a href="GetPass.asp"><font color="#FF3300">点击这里</font></a>重新申请找回密码。</td>
          </tr>
        </table>
<%
		Else
%>
		<script type="text/javascript">
		<!--
		function Trim(strSource){return strSource.replace(/^\s*/,'').replace(/\s*$/,'');}
		function CheckForm(){
			if (Trim(document.getElementById("PassWord").value) == "") {
				alert("请输入新密码！");
				document.getElementById("PassWord").focus();
				return false;
			}
			if (Trim(document.getElementById("PassWord").value).length < 6 ) {
				alert("密码长度至少6位以上！");
				document.getElementById("PassWord").focus();
				return false;
			}
			if (Trim(document.getElementById("RePassWord").value) == "") {
				alert("请输入确认新密码！");
				document.getElementById("RePassWord").focus();
				return false;
			}
			if (Trim(document.getElementById("PassWord").value) != Trim(document.getElementById("RePassWord").value)) {
				alert("新密码与确认新密码不一致！\n\n请重新输入！");
				document.getElementById("RePassWord").focus();
				return false;
			}	
		}
		-->
		</script>
		<form name="LoginForm" action="" method="post" onsubmit="return CheckForm()">
			<input name="PassKey" type="hidden" maxlength="50" value="<%=rs("PassKey")%>" />
			<table width="100%" border="0" cellspacing="5" cellpadding="0">
              <tr>
                <td height="30">&nbsp;</td>
                <td height="30"><font color="#999999">为防止密码被盗，请使用字母、数字及特殊字符组成密码</font></td>
              </tr>
              <tr>
                <td width="100" height="30" align="right">新密码：</td>
                <td><input name="PassWord" type="password" class="ipt1" id="PassWord" value="" maxlength="50"></td>
              </tr>
              <tr>
                <td height="30" align="right">重复新密码：</td>
                <td><input name="RePassWord" type="password" class="ipt1" id="RePassWord" value="" maxlength="50" /></td>
              </tr>
              <tr>
                <td height="30" align="right"></td>
                <td><input name="Submit" type="submit" class="btn1" value="提交" /></td>
              </tr>
            </table>
		</form>
<%
		End If
	Else
%>
		<table width="100%" border="0" cellspacing="5" cellpadding="0">
          <tr>
            <td style="font-size:14px; line-height:200%;"><%=Tip%></td>
          </tr>
        </table>
<%
	End IF
%>

</div>
<%
	CloseConn
%>
</body>
</html>
