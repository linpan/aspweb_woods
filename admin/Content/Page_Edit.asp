<!--#include file="../../Include/Class_Conn.asp"-->
<!--#include file="../../Include/Class_Main.asp"-->
<!--#include file="../p8_Check.asp"-->
<%
	If Request("Submit")<>"" Then
		Dim i,ClassID,Content,Pic
		ClassID    = Trim(Request.Form("ClassID"))
		ClassName  = Trim(Request.Form("ClassName"))
		Pic        = Trim(Request.Form("Pic"))
		Content    = Trim(Request.Form("Content"))

		If ClassID="" Then
			Response.Write "<script>alert(""分类ID丢失！"");window.history.back();</script>"
			Response.End()
		End If

		Set rs=Server.CreateObject("Adodb.Recordset")
		rs.Open "Select Pic,Content From p8_Page Where ClassID = " & ClassID,Conn,1,3
		
		rs("Pic")     = Pic
		rs("Content") = Content
		
		rs.Update
		rs.Close
		Set rs=Nothing
		CloseConn
		Response.Redirect "Page_Edit.asp?Tip=修改成功！&ClassID="& ClassID & "&ClassName="& ClassName
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
<title>修改单页</title>
<script type="text/javascript">top.window.aTitle.innerText='修改单页'</script>
<link href="../css/Public.css" rel="stylesheet" type="text/css" /> 
<script type="text/javascript" src="../Include/TipBox.js"></script>
<script type="text/javascript" src="../Include/calendar.js"></script>
<script type="text/javascript" src="../Include/Pub.js"></script>
<script type="text/javascript">
function doChange(objText, objDrop){
	if (!objDrop) return;
	var str = objText.value;
	var arr = str.split("|");
	objDrop.length=0;
	for (var i=0; i<arr.length; i++){
		objDrop.options[i] = new Option(arr[i], arr[i]);
	}
}
</script>
</head>

<body>
<%
	Dim Tip
	Tip = Request.QueryString("Tip")
	If Tip <> "" Then
		Response.Write "<script type=""text/javascript"">window.onload=function(){new x.creat(1, 41, 5, 10, '"& Tip &"');}</script>"
	End If
	
	'-----------------------------------------------------------------------
	ClassID   = Request("ClassID")
	ClassName = Request("ClassName")
	
	Set rs = Server.CreateObject ("ADODB.Recordset")
	rs.Open "Select id From p8_Page Where ClassID = "& ClassID,Conn,1,1
	
	If rs.Eof Then
		Conn.Execute = "Insert Into p8_Page (ClassID) Values ("& ClassID &")"
	End If
	
	Set rs = Server.CreateObject ("ADODB.Recordset")
	rs.Open "Select Content,Pic From p8_Page Where ClassID = "& ClassID,Conn,1,1
%>
	<form id="AddFrom" name="AddFrom" method="post" action="Page_Edit.asp">
	<input name="ClassID" type="hidden" value="<%=ClassID%>">
	<input name="ClassName" type="hidden" value="<%=ClassName%>">
      <table width="99%" border="0" cellpadding="5" cellspacing="0" style="margin:5px 0;">
        <tr>
          <td style="padding-right:20px;"><strong><%=ClassName%></strong><input name="Pic" type="text" class="ipt4" id="Pic" style="width:300px; display:none;"  readonly="readonly" value="<%=rs("Pic")%>"></td>
        </tr>
        <tr>
          <td width="1150" height="30">
		 	  <textarea name="Content" style="display:none"><%=rs("Content")%></textarea>
              <IFRAME ID="eWebEditor1" src="../Include/Editer/ewebeditor.htm?id=Content&style=coolblue&savepathfilename=Pic" frameborder="0" scrolling="no" width="100%" height="530"></IFRAME></td>
        </tr>
        <tr>
          <td height="30" valign="bottom" style="padding:10px;"><input name="Submit" type="submit" class="btn2" value=" 修 改 " ></td>
        </tr>
      </table>
</form>
<%
	CloseConn
%>
</body>
</html>