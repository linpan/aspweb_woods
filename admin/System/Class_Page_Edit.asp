<!--#include file="../../Include/Class_Conn.asp"-->
<!--#include file="../../Include/Class_Main.asp"-->
<% Super=1 'ֻ����������Ա���ʸ�ҳ %>
<!--#include file="../p8_Check.asp"-->
<%
	Dim ClassType,ClassName,UserLimit,rs
	
	If Request("ClassName") <> "" Then
		id        = Trim(Request("id"))
		ParentID  = Trim(Request("ParentID"))
		ClassName = Trim(Request("ClassName"))
		UserLimit = Trim(Request("UserLimit"))

		If ParentID="" Then
			Response.Write "<script>alert(""�ϼ�ID������ʧ"");window.history.back();</script>"
			Response.End()
		End If

		If ClassName="" Then
			Response.Write "<script>alert(""����д��������"");window.history.back();</script>"
			Response.End()
		End If
		
		If UserLimit="" Then
			Response.Write "<script>alert(""��ѡ���Ķ�Ȩ��"");window.history.back();</script>"
			Response.End()
		End If
		
		Set rs=Server.CreateObject("Adodb.Recordset")
		rs.Open "Select ClassName,ParentID,UserLimit From p8_Class Where id=" & id,Conn,1,3

		rs("ParentID")   = ParentID
		rs("ClassName")  = ClassName
		rs("UserLimit")  = UserLimit

		rs.Update
		rs.Close
		Set rs=Nothing
		CloseConn
		Response.Redirect "Class_Page_List.asp?Tip=�޸ĳɹ���"
		Response.End()
		
	End If
	
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��Ŀ����</title>
<script type="text/javascript">top.window.aTitle.innerText='��Ŀ����'</script>
<meta http-equiv="pragma" content="no-cache" />
<meta http-equiv="Cache-Control" content="no-cache, must-revalidate" />
<meta http-equiv="expires" content="Wed, 26 Feb 1997 08:21:57 GMT" />
<meta http-equiv="expires" content="0" />
<link href="../css/Public.css" rel="stylesheet" type="text/css" /> 
<script type="text/javascript" src="../Include/Pub.js"></script>
<script type="text/javascript" src="../Include/TipBox.js"></script>
<script type="text/javascript">
function Trim(strSource){return strSource.replace(/^\s*/,'').replace(/\s*$/,'');}
function CheckPageForm(){
    if (Trim(document.getElementById("ClassName").value)=="") {
        alert("����д��������");
		document.getElementById("ClassName").focus();
		return false;
    }
}	
</script>
</head>

<body>
<%
	Dim id,rs_pre
	id = Request("id")

	Set rs = Server.CreateObject("ADODB.RecordSet")
	rs.open "Select id,Num,ClassType,ClassLevel,ClassName,ParentID,UserLimit From p8_Class Where id="& id,Conn,1,1
	
	ClassLevel = rs("ClassLevel") '���༶��
	ClassName  = rs("ClassName")  '��������
	ParentID   = rs("ParentID")   '�ϼ�ID
	UserLimit  = rs("UserLimit")  '�Ķ�Ȩ��
	
	If ParentID   = "" Then ParentID   = 0 '�����һ�����࣬���ϼ�ID�Զ�Ϊ0
	If ClassLevel = "" Then ClassLevel = 0 '�����һ�����࣬���ϼ������Զ�Ϊ0
%>
<form name="PageForm" method="post" action="" onSubmit="return CheckPageForm();">
<input name="id" type="hidden" value="<%=id%>">
  <table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
    <tr>
      <td height="25" colspan="2" bgcolor="#E4EDF9"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td>&nbsp;<span class="f14 cBlack">��ҳ - �޸ĵ�ҳ</span></td>
          <td align="right"><a href="javascript:history.back();">&lt;&lt;����</a>&nbsp;&nbsp;</td>
        </tr>
      </table></td>
    </tr>
    <tr>
      <td width="74" height="30" align="right" bgcolor="#F8FBFE">��&nbsp;��&nbsp;����</td>
      <td bgcolor="#F8FBFE">
	  <input name="ParentID" type="hidden" value="0">
	  <input name="ClassName" type="text" class="ipt4" id="ClassName" style="width:150px;" value="<%=ClassName%>" maxlength="80"></td>
    </tr>
    <tr>
      <td height="30" align="right" bgcolor="#F8FBFE">�Ķ�Ȩ�ޣ�</td>
      <td bgcolor="#F8FBFE" class="cBlack"><input name="UserLimit" type="radio" value="0" <%If UserLimit = "0" Then Response.Write "checked"%>>����
      <input type="radio" name="UserLimit" value="1" <%If UserLimit = "1" Then Response.Write "checked"%>>��ע���Ա</td>
    </tr>
    <tr>
      <td height="30" align="center" bgcolor="#F8FBFE">&nbsp;</td>
      <td bgcolor="#F8FBFE"><input name="Submit" type="submit" class="btn2" value=" �� �� " ></td>
    </tr>
  </table>
</form>
</body>
</html>
<%
	CloseConn
%>