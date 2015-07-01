<!--#include file="../../Include/Class_Conn.asp"-->
<!--#include file="../../Include/Class_Main.asp"-->
<% Super=1 '只允许超级管理员访问该页 %>
<!--#include file="../p8_Check.asp"-->
<%
	Dim FieldName,MaxLen,Width,Height,Content,Options,rs
	
	If Request("FieldName") <> "" Then
		id         = Trim(Request("id"))
		FieldName  = Trim(Request("FieldName"))
		MaxLen     = Trim(Request("MaxLen"))
		Width      = Trim(Request("Width"))
		Height     = Trim(Request("Height"))
		Content    = Trim(Request("Content"))
		Options    = Trim(Request("Options"))

		If FieldName="" Then
			Response.Write "<script>alert(""请填写名称"");window.history.back();</script>"
			Response.End()
		End If
		
		Set rs=Server.CreateObject("Adodb.Recordset")
		rs.Open "Select FieldName,MaxLen,FieldType,Width,Height,Content,Options From p8_Field Where id=" & id,Conn,1,3

		rs("FieldName")  = FieldName
		If MaxLen <> "" Then rs("MaxLen") = MaxLen
		If Width <> ""  Then rs("Width")  = Width
		If Height <> "" Then rs("Height") = Height
		rs("Content")    = Content
		rs("Options")    = Options
		
		FieldType = rs("FieldType")

		rs.Update
		rs.Close
		Set rs=Nothing
		
		CloseConn
		Response.Write "<script>alert(""修改成功"");window.close();window.opener.document.getElementById(""Fi"& id &""").innerHTML = ""<strong style='color:#009900;'>"& FieldName &"</strong>("& FieldType &")&nbsp;&nbsp;<a href='#' onclick=\""openw('../System/Class_Form_Field_Edit.asp?Fiid="& id &"','name1',800,500)\"">修改</a>&nbsp;<a href='#' onclick='DelField("& id &")'>删除</a>"";</script>"
		Response.End()
		
	End If
	
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>修改字段</title>
<meta http-equiv="pragma" content="no-cache" />
<meta http-equiv="Cache-Control" content="no-cache, must-revalidate" />
<meta http-equiv="expires" content="Wed, 26 Feb 1997 08:21:57 GMT" />
<meta http-equiv="expires" content="0" />
<link href="../css/Public.css" rel="stylesheet" type="text/css" /> 
<script type="text/javascript" src="../include/Pub.js"></script>
<script type="text/javascript">
	function Trim(strSource){return strSource.replace(/^\s*/,'').replace(/\s*$/,'');}
	
	function Select(FieldType){
		if(Trim(FieldType) == "text"){
			$Get("cMaxLen").style.display  = "";
			$Get("cWidth").style.display   = "";
			$Get("cHeight").style.display  = "none";
			$Get("cContent").style.display = "";
			$Get("cOptions").style.display = "none";
		}
		if(FieldType == "textarea"){
			$Get("cMaxLen").style.display  = "none";
			$Get("cWidth").style.display   = "";
			$Get("cHeight").style.display  = "";
			$Get("cContent").style.display = "";
			$Get("cOptions").style.display = "none";
		}
		if(FieldType == "radio"){
			$Get("cMaxLen").style.display  = "none";
			$Get("cWidth").style.display   = "none";
			$Get("cHeight").style.display  = "none";
			$Get("cContent").style.display = "none";
			$Get("cOptions").style.display = "";
		}
		if(FieldType == "checkbox"){
			$Get("cMaxLen").style.display  = "none";
			$Get("cWidth").style.display   = "none";
			$Get("cHeight").style.display  = "none";
			$Get("cContent").style.display = "none";
			$Get("cOptions").style.display = "";
		}
		if(FieldType == "select"){
			$Get("cMaxLen").style.display  = "none";
			$Get("cWidth").style.display   = "";
			$Get("cHeight").style.display  = "none";
			$Get("cContent").style.display = "none";
			$Get("cOptions").style.display = "";
		}
	}
	

function CheckForm(){
    if (Trim($Get("FieldName").value)=="") {
        alert("请填写名称");
		$Get("FieldName").focus();
		return false;
    }
	if (Trim($Get("FieldType").value)=="") {
        alert("请选择类型");
		$Get("FieldType").focus();
		return false;
    }
}	
</script>
</head>

<body>
<%
	Dim id
	id  = Request("Fiid")

	Set rs = Server.CreateObject("ADODB.RecordSet")
	rs.open "Select * From p8_Field Where id = "& id,Conn,1,1
%>
<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
<form name="FieldForm" method="post" action="" onSubmit="return CheckForm();">
<input name="id" type="hidden" value="<%=id%>">
  <tr>
    <td height="25" colspan="2" bgcolor="#E4EDF9">&nbsp;<span class="f14 cBlack">表单 - 修改字段</span></td>
  </tr>
  <tr>
    <td width="100" height="30" align="right" bgcolor="#F8FBFE">名称：<span class="cYellow">*</span></td>
    <td bgcolor="#F8FBFE" class="cBlack"><input name="FieldName" type="text" class="ipt4" id="FieldName" style="width:150px;" value="<%=rs("FieldName")%>" maxlength="50"></td>
  </tr>
  <tr>
    <td height="30" align="right" bgcolor="#F8FBFE">变量名：<span class="cYellow">*</span></td>
    <td bgcolor="#F8FBFE"><%=rs("Variable")%></td>
  </tr>
  <tr>
    <td height="30" align="right" bgcolor="#F8FBFE">类型：<span class="cYellow">*</span></td>
    <td bgcolor="#F8FBFE" class="cBlack">
	<%=rs("FieldType")%>
	<script type="text/javascript">
		window.onload=function(){Select("<%=rs("FieldType")%>");}
	</script>
	</td>
  </tr>
  <tr id="cMaxLen" style="display:none;">
    <td height="30" align="right" bgcolor="#F8FBFE">内容最大长度：&nbsp;</td>
    <td bgcolor="#F8FBFE"><span class="cBlack">
      <input name="MaxLen" id="MaxLen" type="text" class="ipt4" style="width:150px;" value="<%=rs("MaxLen")%>" maxlength="80" onKeyPress="if((event.keyCode<48 || event.keyCode>57) && event.keyCode!=46 || /\d\d\d\d\d$/.test(value))event.returnValue=false" onMouseDown="this.oncontextmenu = function() { return false;} " onKeyDown="if(event.ctrlKey) return false">
    </span></td>
  </tr>
  <tr id="cWidth" style="display:none;">
    <td height="30" align="right" bgcolor="#F8FBFE">表单宽度：&nbsp;</td>
    <td bgcolor="#F8FBFE"><span class="cBlack">
      <input name="Width" id="Width" type="text" class="ipt4" style="width:150px;" value="<%=rs("Width")%>" maxlength="80" onKeyPress="if((event.keyCode<48 || event.keyCode>57) && event.keyCode!=46 || /\d\d\d\d\d$/.test(value))event.returnValue=false" onMouseDown="this.oncontextmenu = function() { return false;} " onKeyDown="if(event.ctrlKey) return false">
    px</span></td>
  </tr>
  <tr id="cHeight" style="display:none;">
    <td height="30" align="right" bgcolor="#F8FBFE">表单高度：&nbsp;</td>
    <td bgcolor="#F8FBFE"><span class="cBlack">
      <input name="Height" id="Height" type="text" class="ipt4" style="width:150px;" value="<%=rs("Height")%>" maxlength="80" onKeyPress="if((event.keyCode<48 || event.keyCode>57) && event.keyCode!=46 || /\d\d\d\d\d$/.test(value))event.returnValue=false" onMouseDown="this.oncontextmenu = function() { return false;} " onKeyDown="if(event.ctrlKey) return false">
      px</span></td>
  </tr>
  <tr id="cContent" style="display:none;">
    <td height="30" align="right" bgcolor="#F8FBFE">默认内容：&nbsp;</td>
    <td bgcolor="#F8FBFE"><span class="cBlack">
      <input name="Content" id="Content" type="text" class="ipt4" style="width:150px;" value="<%=rs("Content")%>" maxlength="80">
    </span></td>
  </tr>
  <tr id="cOptions" style="display:none;">
    <td height="30" align="right" bgcolor="#F8FBFE">可选内容：&nbsp;</td>
    <td bgcolor="#F8FBFE"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td valign="top"><textarea name="Options" id="Options" class="ipt3" style="width:300px; height:100px;"><%=rs("Options")%></textarea></td>
          <td style="padding-left:10px; line-height:150%; color:#999;">一行为一个选项，如: <BR>
            光电鼠标<BR>
            机械鼠标<BR>
            没有鼠标<BR>
            注意: 字段确定后请勿修改索引和内容的对应关系，但仍可以新增字段。如需调换显示顺序，可以通过移动整行的上下位置来实现</td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td height="30" align="center" bgcolor="#F8FBFE">&nbsp;</td>
    <td bgcolor="#F8FBFE"><input name="Submit" type="submit" class="btn2" value=" 修 改 "></td>
  </tr>
</form>
</table>
</body>
</html>
<%
	rs.Close
	Set rs = Nothing
	CloseConn
%>