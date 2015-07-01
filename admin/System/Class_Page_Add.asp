<!--#include file="../../Include/Class_Conn.asp"-->
<!--#include file="../../Include/Class_Main.asp"-->
<% Super=1 '只允许超级管理员访问该页 %>
<!--#include file="../p8_Check.asp"-->
<%
	Dim Num,ClassType,ClassName,UserLimit,ParentID,ClassLevel,rs
	
	If Request("ClassName") <> "" Then
		Num        = Trim(Request("Num"))
		ClassType  = Trim(Request("ClassType"))
		ClassLevel = Trim(Request("ClassLevel"))
		ClassName  = Trim(Request("ClassName"))
		UserLimit  = Trim(Request("UserLimit"))
		ParentID   = Trim(Request("ParentID"))
		
		If Num="" Then
			Response.Write "<script>alert(""标识符丢失，请返回刷新！"");window.history.back();</script>"
			Response.End()
		End If
		
		If ClassType="" Then
			Response.Write "<script>alert(""上级分类错误"");window.history.back();</script>"
			Response.End()
		End If
		
		If ParentID="" Then
			Response.Write "<script>alert(""上级ID参数丢失"");window.history.back();</script>"
			Response.End()
		End If
		
		If ClassLevel="" Then
			Response.Write "<script>alert(""分类级别参数丢失"");window.history.back();</script>"
			Response.End()
		End If

		If ClassName="" Then
			Response.Write "<script>alert(""请填写分类名称"");window.history.back();</script>"
			Response.End()
		End If
		
		If UserLimit="" Then
			Response.Write "<script>alert(""请选择阅读权限"");window.history.back();</script>"
			Response.End()
		End If
		
		Set rs=Server.CreateObject("Adodb.Recordset")
		rs.Open "Select Num,ClassType,ClassLevel,ParentID,ClassName,UserLimit From p8_Class",Conn,1,3
		rs.AddNew

		rs("Num")        = Num
		rs("ClassType")  = ClassType
		rs("ClassLevel") = ClassLevel + 1
		rs("ParentID")   = ParentID
		rs("ClassName")  = ClassName
		rs("UserLimit")  = UserLimit

		rs.Update
		rs.Close
		Set rs=Nothing
		CloseConn
		Response.Redirect "Class_Page_List.asp?Tip=添加成功！"
		Response.End()
		
	End If
	
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>栏目管理</title>
<script type="text/javascript">top.window.aTitle.innerText='栏目管理'</script>
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
        alert("请填写分类名称");
		document.getElementById("ClassName").focus();
		return false;
    }
}	
</script>
</head>

<body>
<%

	Num = MakeNum()

%>
<form name="PageForm" method="post" action="" onSubmit="return CheckPageForm();">
<input name="Num" type="hidden" value="<%=Num%>">
  <table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
    <tr>
      <td height="25" colspan="2" bgcolor="#E4EDF9"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td>&nbsp;<span class="f14 cBlack">单页 - 增加单页</span></td>
          <td align="right"><a href="javascript:history.back();">&lt;&lt;返回</a>&nbsp;&nbsp;</td>
        </tr>
      </table></td>
    </tr>
    <tr>
      <td width="74" height="30" align="right" bgcolor="#F8FBFE">分&nbsp;类&nbsp;名：</td>
      <td bgcolor="#F8FBFE">
	  <input name="ParentID" type="hidden" value="0">
        <input name="ClassType" type="hidden" value="单页">
		<input name="ClassLevel" type="hidden" value="0">
	  <input name="ClassName" type="text" class="ipt4" id="ClassName" style="width:150px;" maxlength="80"></td>
    </tr>
    <tr>
      <td height="30" align="right" bgcolor="#F8FBFE">阅读权限：</td>
      <td bgcolor="#F8FBFE" class="cBlack"><input name="UserLimit" type="radio" value="0" checked>公开
      <input type="radio" name="UserLimit" value="1">仅注册会员</td>
    </tr>
    <tr>
      <td height="30" align="center" bgcolor="#F8FBFE">&nbsp;</td>
      <td bgcolor="#F8FBFE"><input name="Submit" type="submit" class="btn2" value=" 增 加 " ></td>
    </tr>
  </table>
</form>
</body>
</html>
<%
	CloseConn
%>