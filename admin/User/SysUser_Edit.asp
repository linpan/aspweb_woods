<!--#include file="../../Include/Class_Conn.asp"-->
<!--#include file="../../Include/Class_Main.asp"-->
<!--#include file="../../Include/Class_MD5.asp"-->
<% Super=1 'ֻ����������Ա���ʸ�ҳ %>
<!--#include file="../p8_Check.asp"-->
<%
	If Request("Submit") <> "" Then
		Dim s_User,s_Pass,s_Name,s_Level
		id       = Request("id")
		Page     = Request("Page")
		Ps_Level = Request("Ps_Level")
		Ps_User  = Request("Ps_User")
		Ps_Name  = Request("Ps_Name")
		s_Pass   = Trim(Request("s_Pass"))
		s_Name   = Trim(Request("s_Name"))
		s_Level  = Trim(Request("s_Level"))
		
		If s_Name = "" Then
			Response.Write "<script>alert(""����д����"& s_Name &""");window.history.back();</script>"
			Response.End()
		End If
		
		If s_Level = "" Then
			Response.Write "<script>alert(""����ѡ��Ȩ��"");window.history.back();</script>"
			Response.End()
		End If
		
		Set rs=Server.CreateObject("Adodb.Recordset")
		rs.Open "Select s_Pass,s_Name,s_Level From p8_Super Where id = "& id &"",Conn,1,3
		
		If Not rs.Eof Then
			If s_Pass<>"" Then rs("s_Pass") = MD5(s_Pass)
			rs("s_Name")  = s_Name
			rs("s_Level") = s_Level
			rs.Update
		End If

		rs.Close
		Set rs=Nothing
		
		CloseConn
		Response.Redirect "SysUser_List.asp?Tip=�޸ĳɹ���&Page="& Page &"&Ps_Level="& Ps_Level &"&Ps_User="& Ps_User &"&Ps_Name="& Ps_Name
		Response.End()
		
	End If
	
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�޸Ĺ���Ա</title>
<script type="text/javascript">top.window.aTitle.innerText='�޸Ĺ���Ա'</script>
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
	if(Trim($p8("s_Pass").value) != Trim($p8("s_RePass").value)){
		alert("�ظ����������������ͬ\n\n�粻�޸ģ��뱣��Ϊ��");
		$p8("s_RePass").focus();
		return false;
	}
	if(Trim($p8("s_Name").value) == ''){
		alert("����д����");
		$p8("s_Name").focus();
		return false;
	}

	var a = document.getElementsByName("s_Level");
	var num=0;
	for (var i=0; i<a.length; i++){
		if(a[i].checked) {
			num++;
		}
	}
	if(num==0) {
		alert("��ѡ��Ȩ��");
		return false;
	}
	
	return true;
}
</script>
</head>

<body>
<%
	Dim id,Page,Ps_Level,Ps_User,Ps_Name
	id       = Request("id")
	Page     = Request("Page")
	Ps_Level = Request("s_Level")
	Ps_User  = Request("s_User")
	Ps_Name  = Request("s_Name")
	
	Set rs = Server.Createobject("Adodb.RecordSet")
	rs.open "Select * From p8_Super Where id= " & id ,Conn,1,3
%>
<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
<form name="AddForm" method="post" action="SysUser_Edit.asp" onSubmit="return CheckForm()">
<input type="hidden" name="id" value="<%=id%>" />
<input type="hidden" name="Page" value="<%=Page%>" />
<input type="hidden" name="Ps_Level" value="<%=Ps_Level%>" />
<input type="hidden" name="Ps_User" value="<%=Ps_User%>" />
<input type="hidden" name="Ps_Name" value="<%=Ps_Name%>" />
  <tr>
    <td width="80" height="30" align="right" bgcolor="#F8FBFE">�û�����</td>
    <td bgcolor="#F8FBFE"><%=rs("s_User")%></td>
  </tr>
  <tr>
    <td height="30" align="right" bgcolor="#F8FBFE">���룺</td>
    <td bgcolor="#F8FBFE"><input name="s_Pass" type="password" class="ipt3" id="s_Pass" maxlength="50" style="width:200px;">
      &nbsp;<span class="cGray">�粻�޸ģ��뱣��Ϊ��</span></td>
  </tr>
  <tr>
    <td height="30" align="right" bgcolor="#F8FBFE">�ظ����룺</td>
    <td bgcolor="#F8FBFE"><input name="s_RePass" type="password" class="ipt3" id="s_RePass" maxlength="50" style="width:200px;"></td>
  </tr>
  <tr>
    <td height="30" align="right" bgcolor="#F8FBFE">������</td>
    <td bgcolor="#F8FBFE"><input name="s_Name" type="text" class="ipt3" id="s_Name" maxlength="10" style="width:200px;" value="<%=rs("s_Name")%>"></td>
  </tr>
  <tr>
    <td height="30" align="right" bgcolor="#F8FBFE">Ȩ�ޣ�</td>
    <td bgcolor="#F8FBFE">
	  <input type="radio" name="s_Level" value="1" <%If rs("s_Level")="1" Then Response.Write " checked=""checked"""%>>��������Ա
      <input type="radio" name="s_Level" value="2" <%If rs("s_Level")="2" Then Response.Write " checked=""checked"""%>>¼��Ա</td>
  </tr>
  <tr>
    <td height="30" align="right" bgcolor="#F8FBFE">&nbsp;</td>
    <td bgcolor="#F8FBFE"><span style="padding:20px 0;">
      <input name="Submit" type="submit" class="btn2" value=" �޸� " >
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
	CloseRs
	CloseConn
%>