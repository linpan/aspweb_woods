<!--#include file="../../Include/Class_Conn.asp"-->
<!--#include file="../../Include/Class_Main.asp"-->
<% Super=1 'ֻ����������Ա���ʸ�ҳ %>
<!--#include file="../p8_Check.asp"-->
<%
	Dim Num,FieldName,Variable,FieldType,MaxLen,Width,Height,Content,Options,rs
	
	If Request("FieldName") <> "" Then
		Num        = Trim(Request("Num"))
		FieldName  = Trim(Request("FieldName"))
		Variable   = Trim(Request("Variable"))
		FieldType  = Trim(Request("FieldType"))
		MaxLen     = Trim(Request("MaxLen"))
		Width      = Trim(Request("Width"))
		Height     = Trim(Request("Height"))
		Content    = Trim(Request("Content"))
		Options    = Trim(Request("Options"))
		
		If Num="" Then
			Response.Write "<script>alert(""��ʶ����ʧ���뷵��ˢ�£�"");window.history.back();</script>"
			Response.End()
		End If
		
		If FieldName="" Then
			Response.Write "<script>alert(""����д����"");window.history.back();</script>"
			Response.End()
		End If
		
		If Variable="" Then
			Response.Write "<script>alert(""����д������"");window.history.back();</script>"
			Response.End()
		End If
		
		If FieldType="" Then
			Response.Write "<script>alert(""��ѡ������"");window.history.back();</script>"
			Response.End()
		End If
		
		Set rs2=Server.CreateObject("Adodb.Recordset")
		rs2.Open "Select id From p8_Field Where ClassNum = '"& ClassNum &"' And Variable = '"& Variable &"'",Conn,1,1
			If Not rs2.Eof Then
				Response.Write "<script>alert(""�������Ѵ��ڣ����޸ĺ��ύ��"");window.history.back();</script>"
				Response.End()
			End If
		rs2.Close
		Set rs2 = Nothing
		
		Set rs=Server.CreateObject("Adodb.Recordset")
		rs.Open "Select FieldName,Variable,FieldType,MaxLen,Width,Height,Content,Options,ClassNum From p8_Field",Conn,1,3
		rs.AddNew

		rs("FieldName")  = FieldName
		rs("Variable")   = Variable
		rs("FieldType")  = FieldType
		If MaxLen <> "" Then rs("MaxLen") = MaxLen
		If Width <> ""  Then rs("Width")  = Width
		If Height <> "" Then rs("Height") = Height
		rs("Content")    = Content
		rs("Options")    = Options
		rs("ClassNum")   = Num

		rs.Update
		rs.Close
		Set rs=Nothing
		
		'��ȡ����ID
		Dim Fiid
		Set rs=Server.CreateObject("Adodb.Recordset")
		rs.Open "Select Top 1 id From p8_Field Order By id Desc",Conn,1,1
			Fiid = rs("id")
		rs.Close
		Set rs=Nothing
		
		CloseConn
		Response.Write "<script>alert(""��ӳɹ�"");window.close();window.opener.document.getElementById(""Field"").innerHTML = ""<div style='padding:5px 0;' id='Fi"& Fiid &"'><strong style='color:#009900;'>"& FieldName &"</strong>("& FieldType &")&nbsp;&nbsp;<a href='#' onclick=\""openw('../System/Class_Form_Field_Edit.asp?Fiid="& Fiid &"','name1',800,500)\"">�޸�</a>&nbsp;<a href='#' onclick='DelField("& Fiid &")'>ɾ��</a></div>"" + window.opener.document.getElementById(""Field"").innerHTML;</script>"
		Response.End()
		
	End If
	
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>����ֶ�</title>
<meta http-equiv="pragma" content="no-cache" />
<meta http-equiv="Cache-Control" content="no-cache, must-revalidate" />
<meta http-equiv="expires" content="Wed, 26 Feb 1997 08:21:57 GMT" />
<meta http-equiv="expires" content="0" />
<link href="../css/Public.css" rel="stylesheet" type="text/css" /> 
<script type="text/javascript" src="../include/Pub.js"></script>
<script type="text/javascript">
	function Select(FieldType){
		if(FieldType == "text"){
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
	
function Trim(strSource){return strSource.replace(/^\s*/,'').replace(/\s*$/,'');}
function CheckForm(){
    if (Trim($Get("FieldName").value)=="") {
        alert("����д����");
		$Get("FieldName").focus();
		return false;
    }
	if (Trim($Get("Variable").value)=="") {
        alert("����д������");
		$Get("Variable").focus();
		return false;
    }
	if (Trim($Get("Variable").value)!=""){
	  //��ǩ����ֻ�ܰ���Ӣ����ĸ������,��ֻ������ĸ��ͷ
	  var re = /^([a-zA-Z]([a-zA-Z0-9])*)$/igm;
	  var SortDirName = $Get("Variable").value;
	  if(re.test(SortDirName) == false){
		alert("������ֻ������ĸ��ͷ��");
		$Get("Variable").focus();
		return false;
	  }      
	}
	if ($Get("Tip_1").innerHTML=="�������Ѵ��ڣ������������������") {
        alert("�������Ѵ��ڣ����޸ĺ��ύ��");
		$Get("Variable").focus();
		return false;
    }
	if (Trim($Get("FieldType").value)=="") {
        alert("��ѡ������");
		$Get("FieldType").focus();
		return false;
    }
}	
//xmlHttp
var xmlHttp = false;
try{xmlHttp = new ActiveXObject("Msxml2.XMLHTTP");}
catch (e){
	try{xmlHttp = new ActiveXObject("Microsoft.XMLHTTP");}
	catch(e2){xmlHttp=false;}
}
if (!xmlHttp && typeof XMLHttpRequest != 'undefined'){xmlHttp = new XMLHttpRequest();}

//<���������Ƿ�ռ��>
function VariableUse(){	
	if($Get("Variable").value){
		var Variable=$Get("Variable").value;
		var VariableUrl="Class_Form_Field_Check.asp?ClassNum=<%=Request("Num")%>&Variable="+Variable;
		xmlHttp.open("GET",VariableUrl,true);
		xmlHttp.onreadystatechange=UpdateVariable;
		xmlHttp.send(null);
	}
}

function UpdateVariable(){
	if (xmlHttp.readyState==3){
		$Get("Tip_1").innerHTML = "";
	}
	if (xmlHttp.readyState==4){
		var uback=xmlHttp.responseText;
		if(uback=="1"){
			$Get("Tip_1").innerHTML = "<font color='red'>�������Ѵ��ڣ������������������</font>";
		}else{
			$Get("Tip_1").innerHTML = "<font color='green'>����ʹ��</font>";
		}
	}
}
//</���������Ƿ�ռ��>
</script>
</head>

<body>
<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
<form name="FieldForm" method="post" action="" onSubmit="return CheckForm();">
<input name="Num" type="hidden" value="<%=Request("Num")%>">
  <tr>
    <td height="25" colspan="2" bgcolor="#E4EDF9">&nbsp;<span class="f14 cBlack">�� - ����ֶ�</span></td>
  </tr>
  <tr>
    <td width="100" height="30" align="right" bgcolor="#F8FBFE">���ƣ�<span class="cYellow">*</span></td>
    <td bgcolor="#F8FBFE" class="cBlack"><input name="FieldName" type="text" class="ipt4" id="FieldName" style="width:150px;" value="" maxlength="50"></td>
  </tr>
  <tr>
    <td height="30" align="right" bgcolor="#F8FBFE">��������<span class="cYellow">*</span></td>
    <td bgcolor="#F8FBFE">
      <input name="Variable" type="text" class="ipt4" id="Variable" style="width:150px;" value="" onBlur="VariableUse();" maxlength="50" onKeyUp="value=value.replace(/[^\a-\z\A-\Z\d]/g,'')">
    <span id="Tip_1"></span></td>
  </tr>
  <tr>
    <td height="30" align="right" bgcolor="#F8FBFE">���ͣ�<span class="cYellow">*</span></td>
    <td bgcolor="#F8FBFE" class="cBlack">
	<select name="FieldType" id="FieldType" onChange="Select(this.value)">
      <option value=""></option>
	  <option value="text">�����ı�(text)</option>
      <option value="textarea">�����ı�(textarea)</option>
      <option value="radio">��ѡ��(radio)</option>
      <option value="checkbox">��ѡ��(checkbox)</option>
      <option value="select">�����б�(select)</option>
    </select>    </td>
  </tr>
  <tr id="cMaxLen" style="display:none;">
    <td height="30" align="right" bgcolor="#F8FBFE">������󳤶ȣ�&nbsp;</td>
    <td bgcolor="#F8FBFE"><span class="cBlack">
      <input name="MaxLen" id="MaxLen" type="text" class="ipt4" style="width:150px;" value="" maxlength="80" onKeyPress="if((event.keyCode<48 || event.keyCode>57) && event.keyCode!=46 || /\d\d\d\d\d$/.test(value))event.returnValue=false" onMouseDown="this.oncontextmenu = function() { return false;} " onKeyDown="if(event.ctrlKey) return false">
    </span></td>
  </tr>
  <tr id="cWidth" style="display:none;">
    <td height="30" align="right" bgcolor="#F8FBFE">����ȣ�&nbsp;</td>
    <td bgcolor="#F8FBFE"><span class="cBlack">
      <input name="Width" id="Width" type="text" class="ipt4" style="width:150px;" value="" maxlength="80" onKeyPress="if((event.keyCode<48 || event.keyCode>57) && event.keyCode!=46 || /\d\d\d\d\d$/.test(value))event.returnValue=false" onMouseDown="this.oncontextmenu = function() { return false;} " onKeyDown="if(event.ctrlKey) return false">
    px</span></td>
  </tr>
  <tr id="cHeight" style="display:none;">
    <td height="30" align="right" bgcolor="#F8FBFE">���߶ȣ�&nbsp;</td>
    <td bgcolor="#F8FBFE"><span class="cBlack">
      <input name="Height" id="Height" type="text" class="ipt4" style="width:150px;" value="" maxlength="80" onKeyPress="if((event.keyCode<48 || event.keyCode>57) && event.keyCode!=46 || /\d\d\d\d\d$/.test(value))event.returnValue=false" onMouseDown="this.oncontextmenu = function() { return false;} " onKeyDown="if(event.ctrlKey) return false">
      px</span></td>
  </tr>
  <tr id="cContent" style="display:none;">
    <td height="30" align="right" bgcolor="#F8FBFE">Ĭ�����ݣ�&nbsp;</td>
    <td bgcolor="#F8FBFE"><span class="cBlack">
      <input name="Content" id="Content" type="text" class="ipt4" style="width:150px;" value="" maxlength="80">
    </span></td>
  </tr>
  <tr id="cOptions" style="display:none;">
    <td height="30" align="right" bgcolor="#F8FBFE">��ѡ���ݣ�&nbsp;</td>
    <td bgcolor="#F8FBFE"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td valign="top"><textarea name="Options" id="Options" class="ipt3" style="width:300px; height:100px;"></textarea></td>
          <td style="padding-left:10px; line-height:150%; color:#999;">һ��Ϊһ��ѡ���: <BR>
            ������<BR>
            ��е���<BR>
            û�����<BR>
          ע��: �ֶ�ȷ���������޸����������ݵĶ�Ӧ��ϵ�����Կ��������ֶΡ����������ʾ˳�򣬿���ͨ���ƶ����е�����λ����ʵ��</td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td height="30" align="center" bgcolor="#F8FBFE">&nbsp;</td>
    <td bgcolor="#F8FBFE"><input name="Submit" type="submit" class="btn2" value=" �� �� "></td>
  </tr>
</form>
</table>
</body>
</html>
