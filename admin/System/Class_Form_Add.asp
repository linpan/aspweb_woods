<!--#include file="../../Include/Class_Conn.asp"-->
<!--#include file="../../Include/Class_Main.asp"-->
<% Super=1 'ֻ����������Ա���ʸ�ҳ %>
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
			Response.Write "<script>alert(""��ʶ����ʧ���뷵��ˢ�£�"");window.history.back();</script>"
			Response.End()
		End If
		
		If ClassType="" Then
			Response.Write "<script>alert(""�ϼ��������"");window.history.back();</script>"
			Response.End()
		End If
		
		If ParentID="" Then
			Response.Write "<script>alert(""�ϼ�ID������ʧ"");window.history.back();</script>"
			Response.End()
		End If
		
		If ClassLevel="" Then
			Response.Write "<script>alert(""���༶�������ʧ"");window.history.back();</script>"
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
		Response.Redirect "Class_Form_List.asp?Tip=��ӳɹ���"
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
function CheckFormForm(){
    if (Trim(document.getElementById("ClassName").value)=="") {
        alert("����д��������");
		document.getElementById("ClassName").focus();
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

//<ɾ���ֶ�>
function DelField(id){	
	obj = $Get("Fi"+ id +"");
	if(id){
		var DelUrl="Class_Form_Field_Del.asp?id="+id;
		xmlHttp.open("GET",DelUrl,true);
		xmlHttp.onreadystatechange=UpdateDel;
		xmlHttp.send(null);
	}
	
	function UpdateDel(){
		if (xmlHttp.readyState==3){}
		if (xmlHttp.readyState==4){
			var uback=xmlHttp.responseText;
			if(uback=="1"){
				obj.style.display = "none";
				new x.creat(1, 41, 5, 10, 'ɾ���ɹ�');
			}
		}
	}
}
//</ɾ���ֶ�>
</script>
</head>

<body>
<%

	Num = MakeNum()

%>
<form name="FormForm" method="post" action="" onSubmit="return CheckFormForm();">
<input name="Num" type="hidden" value="<%=Num%>">
  <table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
    <tr>
      <td height="25" colspan="2" bgcolor="#E4EDF9"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td>&nbsp;<span class="f14 cBlack">�� - ���ӱ�</span></td>
          <td align="right"><a href="javascript:history.back();">&lt;&lt;����</a>&nbsp;&nbsp;</td>
        </tr>
      </table></td>
    </tr>
    <tr>
      <td width="74" height="30" align="right" bgcolor="#F8FBFE">��&nbsp;��&nbsp;����</td>
      <td bgcolor="#F8FBFE">
	  <input name="ParentID" type="hidden" value="0">
        <input name="ClassType" type="hidden" value="��">
		<input name="ClassLevel" type="hidden" value="0">
	  <input name="ClassName" type="text" class="ipt4" id="ClassName" style="width:150px;" maxlength="80"></td>
    </tr>
    <tr>
      <td height="30" align="right" bgcolor="#F8FBFE">�Ķ�Ȩ�ޣ�</td>
      <td bgcolor="#F8FBFE" class="cBlack"><input name="UserLimit" type="radio" value="0" checked>����
      <input type="radio" name="UserLimit" value="1">��ע���Ա</td>
    </tr>
    <tr>
      <td height="30" align="center" bgcolor="#F8FBFE">�Զ����ֶΣ�</td>
      <td bgcolor="#F8FBFE">
	  <div id="Field">	  </div>
	  <a href="#" class="AddBtn" onClick="openw('../System/Class_Form_Field.asp?Num=<%=Num%>','name1',800,500)">����ֶ�</a></td>
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