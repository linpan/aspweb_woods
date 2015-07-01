<!--#include file="../Include/Class_Conn.asp"-->
<!--#include file="../Include/Class_Main.asp"-->
<!--#include file="../Include/Class_MD5.asp"-->
<%
	Dim Email,rs,Fso,FileAll,HtmlStr,PassKey
	Email = Replace_Text(Request("Email"))
	
	If Email <> "" Then
	Check_url()
		Set rs = Server.CreateObject ("ADODB.Recordset")
		rs.Open "Select Email,UserName From p8_User Where Email = '"& Email &"'",Conn,1,1
		
		If Not rs.Eof Then
			If Session("GetPass") = "" Then
				
				PassKey =  Genkey(20)
				
				'//发送邮件给会员
				HtmlStr = "亲爱的"& rs("UserName") &"：<br><br>&nbsp;&nbsp;&nbsp;&nbsp;您好！您正在使用<font color=""blue"">找回密码</font>功能，请通过以下链接设置您的新密码。<br><br>&nbsp;&nbsp;&nbsp;&nbsp;<a href=""http://"& Request.ServerVariables("SERVER_NAME") & SiteDir &"User/SetPass.asp?K="& PassKey &""" target=""_blank"">http://"& Request.ServerVariables("SERVER_NAME") & SiteDir &"User/SetPass.asp?K="& PassKey &"</a>"
				Call SendEmail(Email,HtmlStr,"找回密码邮件")
				
				conn.execute "UpDate p8_User Set PassKey='"& PassKey &"' Where Email='"& Email &"'"  '将PassKey写入会员表
			End If
			
			Tip = "<font color=""red"">密码修改链接已发送到您邮箱，请登录邮箱查收！&nbsp;&nbsp;<a href=""http://"& Mid(Email,Instr(1,Email,"@")+1,Len(Email)) &""" class=""cRed"" target=""_blank"">http://"& Mid(Email,Instr(1,Email,"@")+1,Len(Email)) &"</a></font>"
		Else
			Response.Write "<script>alert(""Emial不存在！"");window.history.back();</script>"
			Response.End()
		End If
		
	End If
%>
<%
Function Genkey(digits)
	Dim char_array(41)
	char_array(0) = "0"
	char_array(1) = "1"
	char_array(2) = "2"
	char_array(3) = "3"
	char_array(4) = "4"
	char_array(5) = "5"
	char_array(6) = "6"
	char_array(7) = "7"
	char_array(8) = "8"
	char_array(9) = "9"
	char_array(10) = "A"
	char_array(11) = "B"
	char_array(12) = "C"
	char_array(13) = "D"
	char_array(14) = "E"
	char_array(15) = "F"
	char_array(16) = "G"
	char_array(17) = "H"
	char_array(18) = "I"
	char_array(19) = "J"
	char_array(20) = "K"
	char_array(21) = "L"
	char_array(22) = "M"
	char_array(23) = "N"
	char_array(24) = "O"
	char_array(25) = "P"
	char_array(26) = "Q"
	char_array(27) = "R"
	char_array(28) = "S"
	char_array(29) = "T"
	char_array(30) = "U"
	char_array(31) = "V"
	char_array(32) = "W"
	char_array(33) = "X"
	char_array(34) = "Y"
	char_array(35) = "Z"
	char_array(36) = "@"
	char_array(37) = ")"
	char_array(38) = "("
	char_array(39) = "_"
	char_array(40) = "$"
	
	Randomize
	
	Do While len(output) < digits
		num = char_array(int((40 - 0 + 1) * rnd + 0))
		output = output + num
	Loop
	
	Genkey = output
End Function
%> 
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>找回密码</title>
<link href="css/Public.css" rel="stylesheet" type="text/css" />
<script type="text/javascript">
function $p8(Obj){return document.getElementById(Obj);}
function Trim(strSource){return strSource.replace(/^\s*/,'').replace(/\s*$p8/,'');}
function CheckForm(){
	if(Trim($p8("Email").value) == ''){
		alert("请填写Email");
		$p8("Email").focus();
		return false;
	}
}
</script>
</head>

<body>
<div style="width:560px; margin-top:50px; padding:60px 20px 20px 20px; background:url(images/Pass_Top.jpg) no-repeat; border:1px solid #1b72af;">
	<%If Tip = "" Then%>
	<table width="100%" border="0" cellspacing="5" cellpadding="0">
		<form method="post" action="" onSubmit="return CheckForm()">
			<tr>
			  <td width="100" height="57" align="right">Email：</td>
			  <td><input name="Email" type="text" class="ipt1" id="Email" value="" maxlength="50" ></td></tr>
			<tr>
              <td height="30" align="right"></td>
			  <td><input name="Submit" type="submit" class="btn1" value="提交" /></td>
		  </tr>
			<tr>
			  <td height="30" align="right"></td>
			  <td align="right" valign="bottom"><a href="<%=SiteDir%>"><strong>返回首页&gt;&gt;</strong></a></td>
		  </tr>
		</form>
	</table>
	<%Else%>
    <table width="100%" border="0" cellspacing="5" cellpadding="0">
      <form method="post" action="" onsubmit="return CheckForm()">
        <tr>
          <td height="100" colspan="2" align="center" style="font-size:14px;"><%=Tip%></td>
        </tr>
        
        <tr>
          <td height="30" align="right"></td>
          <td align="right" valign="bottom"><a href="<%=SiteDir%>"><strong>返回首页&gt;&gt;</strong></a></td>
        </tr>
      </form>
    </table>
	<%End If%>
</div>
<%
	CloseConn
%>
</body>
</html>
