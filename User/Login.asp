<!--#include file="../Include/Class_Conn.asp"-->
<!--#include file="../Include/Class_Main.asp"-->
<!--#include file="../Include/Class_MD5.asp"-->
<%
	Dim UserName,PassWord,rs
	UserName = Replace_Text(Request.Form("UserName"))
	PassWord = Request.Form("PassWord")

	If UserName<>"" And PassWord<>"" Then

		Check_url()	
		Set rs = Server.CreateObject ("ADODB.Recordset")
		rs.Open "Select UserName,PassWord,LoginIP,LoginCount,LoginDate,LoginErr,LoginErrDate,UserState From p8_User Where UserName='"& UserName &"'",Conn,1,1

		If Not rs.Eof Then
			
			If DateDiff("n",rs("LoginErrDate"),Now)>UserLgLock Then Conn.Execute = "Update p8_User Set LoginErr=0 Where UserName='"& UserName &"'" '�����½�������x���ӣ�����������0
			
			If rs("LoginErr")>=UserLgErr And DateDiff("n",rs("LoginErrDate"),Now)<=UserLgLock Then '�����½����x�����ϣ���x�����ڲ������½
				Response.Write "<script>alert(""����¼����������࣬���Ժ��ԣ�"");window.history.back();</script>"
				Response.End()
			Else
				If rs("PassWord") <> MD5(PassWord) Then
					conn.Execute = "Update p8_User Set LoginErr=LoginErr+1,LoginErrDate='"& Now() &"' Where UserName='"& UserName &"'" '�ۼƴ������
					Response.Write "<script>alert(""�û������������"");window.history.back();</script>"
					Response.End()
				Else					
					Session("UserName")  = rs("UserName")
					Session("UserState") = rs("UserState")
					
					Conn.Execute = "Update p8_User Set LoginDate='"& Now() &"',LoginIP='"& Request.Servervariables("Remote_Addr") &"',LoginCount=LoginCount+1,LoginErr=0 Where UserName='"& UserName &"'" '�����û�����¼��Ϣ
					
					Response.Redirect "Main.asp"
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
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��Ա��¼</title>
<link href="css/Public.css" rel="stylesheet" type="text/css" />
<script type="text/javascript">
function $p8(Obj){return document.getElementById(Obj);}
function Trim(strSource){return strSource.replace(/^\s*/,'').replace(/\s*$p8/,'');}
function CheckForm(){
	if(Trim($p8("UserName").value) == ''){
		alert("����д�û���");
		$p8("UserName").focus();
		return false;
	}
	if(Trim($p8("PassWord").value) == ''){
		alert("����д����");
		$p8("PassWord").focus();
		return false;
	}
}
</script>
</head>

<body>
<div style="width:560px; margin-top:50px; padding:60px 20px 20px 20px; background:url(images/Login_Top.jpg) no-repeat; border:1px solid #1b72af;">
	<table width="100%" border="0" cellspacing="5" cellpadding="0">
		<form method="post" action="" onSubmit="return CheckForm()">
			<tr>
			  <td width="100" height="30" align="right">�û���/Email��</td>
			  <td><input name="UserName" type="text" class="ipt1" id="UserName" value="" maxlength="50" ></td></tr>
			<tr>
			  <td height="30" align="right">���룺</td>
			  <td><input name="PassWord" type="password" class="ipt1" id="PassWord" value="" maxlength="50"></td>
			</tr>
			<tr>
              <td height="30" align="right"></td>
			  <td><input name="Submit" type="submit" class="btn1" value="��¼" /></td>
		  </tr>
			<tr>
			  <td height="30" align="right"></td>
			  <td><a href="Register.asp">����ע��</a>&nbsp;&nbsp;<a href="GetPass.asp">�һ�����</a></td>
		  </tr>
		</form>
	</table>
</div>
<%
	CloseConn
%>
</body>
</html>
