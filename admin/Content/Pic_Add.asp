<!--#include file="../../Include/Class_Conn.asp"-->
<!--#include file="../../Include/Class_Main.asp"-->
<!--#include file="../p8_Check.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta http-equiv="pragma" content="no-cache" />
<meta http-equiv="Cache-Control" content="no-cache, must-revalidate" />
<meta http-equiv="expires" content="Wed, 26 Feb 1997 08:21:57 GMT" />
<meta http-equiv="expires" content="0" />
<title>���ͼƬ</title>
<script type="text/javascript">top.window.aTitle.innerText='���ͼƬ'</script>
<link href="../css/Public.css" rel="stylesheet" type="text/css" /> 
<script type="text/javascript" src="../Include/Pub.js"></script>
<script type="text/javascript" src="../Include/TipBox.js"></script>
<script type="text/javascript" src="../Include/calendar.js"></script>
<script type="text/javascript">
function CheckForm(){
    if (document.getElementById("ClassID").value=="") {
        alert("��ѡ�����");
		document.getElementById("ClassID").focus();
		return false;
    }
	if (document.getElementById("PicName").value=="") {
        alert("����д����");
		document.getElementById("PicName").focus();
		return false;
    }
	if (document.getElementById("SmallPic").value=="") {
        alert("���ϴ�����ͼ");
		document.getElementById("SmallPic").focus();
		return false;
    }
}

function doChange(objText, objDrop){
	if (!objDrop) return;
	var str = objText.value;
	var arr = str.split("|");
	objDrop.length=0;
	for (var i=0; i<arr.length; i++){
		objDrop.options[i] = new Option(arr[i], arr[i]);
	}
}

function ShowField(Id){	
	<%
		Set rs = Server.Createobject("Adodb.RecordSet")
		rs.open "Select id From p8_Class Where ClassType='ͼƬ' Order By id Desc",Conn,1,1
			Do While Not rs.Eof
				Response.Write "$Get(""FieldDiv"& rs("id") &""").style.display = ""none"";" & chr(13)
			rs.MoveNext
			Loop
		rs.Close
		Set rs = Nothing
	%>		
	$Get("FieldDiv"+ Id +"").style.display = "block";
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
	
	
	ClassID = Request("ClassID")
%>
	<form id="AddFrom" name="AddFrom" enctype="multipart/form-data" method="post" action="Pic_Add_Sql.asp" onSubmit="return CheckForm();">
      <table width="100%" border="0" cellpadding="0" cellspacing="0" style="margin:5px 0;">
        <tr>
          <td width="82" height="30" align="right">���ࣺ<span class="cRed">*</span>&nbsp;</td>
          <td width="317">
		<select name="ClassID" id="ClassID" style="width:120px;" class="ipt5" onChange="ShowField(this.value)">
		 <option value=""></option>
		  <%
			Set rs = Server.Createobject("Adodb.RecordSet")
			rs.open "Select id,ClassName From p8_Class Where ClassType='ͼƬ' And ClassLevel=1 Order By id Desc",Conn,1,1
			
				Do While Not rs.Eof 
					
					If Clng(ClassID) = Clng(rs("id")) Then
						Response.Write "<option value="""& rs("id") &""" selected=""selected"">"& rs("ClassName") &"</option>"
					Else
						Response.Write "<option value="""& rs("id") &""">"& rs("ClassName") &"</option>"
					End If
					
					Set rsnxt = Server.CreateObject("ADODB.RecordSet")
					rsnxt.open "Select id,ClassName From p8_Class Where ClassType='ͼƬ' And ClassLevel=2 And ParentID="& rs("id") &" Order By id Desc",Conn,1,1
					
					Do While Not rsnxt.Eof 
						Response.Write "<option value="""& rsnxt("id") &""">�� "& rsnxt("ClassName") &"</option>"
						rsnxt.MoveNext         
					Loop 
					
					rsnxt.Close
					Set rsnxt = Nothing

				
				rs.MoveNext         
				Loop 

			rs.Close
			Set rs = Nothing
		  %>
		</select>		</td>
          <td width="769" rowspan="3"><img id="ShowImg" onerror="this.src='../images/NoPhoto.gif'" src="../images/NoPhoto.gif" style="height:75px; border:1px solid #E8E8E8; padding:2px; margin:5px;" /></td>
        </tr>
        <tr>
          <td width="82" height="30" align="right">���ƣ�<span class="cRed">*</span>&nbsp;</td>
          <td>
			<input name="PicName" type="text" class="ipt4" id="PicName" style="width:290px;" maxlength="100"></td>
        </tr>
        <tr>
          <td height="30" align="right">����ͼ��<span class="cRed">*</span>&nbsp;</td>
          <td>
		  <input name="Pic" type="text" class="ipt4" id="Pic" style="width:300px; display:none;"  readonly="readonly">
		  <input name="SmallPic" type="file" class="ipt4" id="SmallPic" style="width:290px;"  value="" size="20" maxlength="50" onChange="$Get('ShowImg').src=this.value" /></td>
        </tr>
        <tr>
          <td width="82" height="30" align="right">���ݣ�<span class="cRed">*</span>&nbsp;</td>
          <td colspan="2" style="padding-right:20px;"><textarea name="Content" style="display:none"></textarea>
              <IFRAME ID="eWebEditor1" src="../Include/Editer/ewebeditor.htm?id=Content&style=coolblue&savepathfilename=Pic" frameborder="0" scrolling="no" width="100%" height="430"></IFRAME></td>
        </tr>
        <tr>
          <td colspan="3" align="right">
<%
	'�Զ����ֶ� ====================================================================================
	Dim ClassNum,FieldStr,FieldId,m

	Set rsClass = Server.Createobject("Adodb.RecordSet")
	rsClass.open "Select id From p8_Class Where ClassType='ͼƬ' Order By id Desc",Conn,1,1 '��ȡ���з���ID,���ݷ���ID�г����з����Զ����ֶ�

	Do While Not rsClass.Eof
	
		Response.Write "<div id=""FieldDiv"& rsClass("id") &""" style=""display:none;""><table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		
		Set rs = Server.Createobject("Adodb.RecordSet")
		rs.open "Select Num From p8_Class Where id="& rsClass("id"),Conn,1,1
		ClassNum = rs("Num")
		rs.Close
		Set rs = Nothing
		
		Set rs = Server.Createobject("Adodb.RecordSet")
		rs.open "Select id,FieldName,Variable,FieldType,MaxLen,Width,Height,Content,Options From p8_Field Where ClassNum='"& ClassNum &"' Order By id Desc",Conn,1,1
		
			FieldStr = ""
			FieldId  = ""
			Fz_FieldId = "Fz_FieldId"& rsClass("id") &"_" '����Ӧ�ý��ܵ��ֶΣ������Ϸ���ID����ָ�����ܣ����ò�ͬ����ͬ���ֶ�

			Do While Not rs.Eof 
				
				FieldId = FieldId & ","& Fz_FieldId & rs("Variable") '����Ӧ�ý��ܵ��ֶΣ������Ϸ����ʶ���з���
				
				If rs("MaxLen") <> 0 Then 
					MaxLen = " maxlength="""& rs("MaxLen") &""""
				End If
				
				'�����ı��� =============================================
				If rs("FieldType") = "text" Then
					FieldStr = FieldStr & "<tr><td width=""82"" height=""30"" align=""right"">"& rs("FieldName") &"��&nbsp;</td><td colspan=""6""><input name="""& Fz_FieldId & rs("Variable") &""" type=""text"" class=""ipt4"" value="""& rs("Content") &""" style=""width:"& rs("Width") &";"" "& MaxLen &"></td></tr>"
				End If
				
				'�����ı��� =============================================
				If rs("FieldType") = "textarea" Then
					FieldStr = FieldStr & "<tr><td width=""82"" height=""30"" align=""right"">"& rs("FieldName") &"��&nbsp;</td><td colspan=""6"" style=""padding-top:5px;""><textarea name="""& Fz_FieldId & rs("Variable") &""" class=""ipt3"" style=""width:"& rs("Width") &"px; height:"& rs("Height") &"px;"">"& rs("Content") &"</textarea></td></tr>"
				End If
				
				'��ѡ�� =================================================
				If rs("FieldType") = "radio" Then
					FieldStr = FieldStr & "<tr><td width=""82"" height=""30"" align=""right"">"& rs("FieldName") &"��&nbsp;</td><td colspan=""6"">"
					
					Options = Split(rs("Options"),chr(13))
					
					For j = 0 To Ubound(Options)
						FieldStr = FieldStr & "<input type=""radio"" name="""& Fz_FieldId & rs("Variable") &""" value=""|"& Replace(Trim(Options(j)),chr(10),"") &"|"" />"& Trim(Options(j)) &" " 
					Next
					
					FieldStr = FieldStr & "</td></tr>"
				End If
				
				'��ѡ�� =================================================
				If rs("FieldType") = "checkbox" Then
					FieldStr = FieldStr & "<tr><td width=""82"" height=""30"" align=""right"">"& rs("FieldName") &"��&nbsp;</td><td colspan=""6"">"
					
					Options = Split(rs("Options"),chr(13))
					
					For j = 0 To Ubound(Options)
						FieldStr = FieldStr & "<input type=""checkbox"" name="""& Fz_FieldId & rs("Variable") &""" value=""|"& Replace(Trim(Options(j)),chr(10),"") &"|"" />"& Trim(Options(j)) &" " 
					Next
					
					FieldStr = FieldStr & "</td></tr>"
				End If
				
				'������ =================================================
				If rs("FieldType") = "select" Then
					FieldStr = FieldStr & "<tr><td width=""82"" height=""30"" align=""right"">"& rs("FieldName") &"��&nbsp;</td><td colspan=""6"">"
					
					Options = Split(rs("Options"),chr(13))
					
					FieldStr = FieldStr & "<select style=""width:"& rs("Width") &";"" name="""& Fz_FieldId & rs("Variable") &"""><option value=""""></option>"
					
					For j = 0 To Ubound(Options)
						FieldStr = FieldStr & "<option value=""|"& Replace(Trim(Options(j)),chr(10),"") &"|"">"& Trim(Options(j)) &"</option>" 
					Next
					
					FieldStr = FieldStr & "</select>"
					
					FieldStr = FieldStr & "</td></tr>"
				End If
			
			rs.MoveNext         
			Loop 
	
		rs.Close
		Set rs = Nothing
		
		If Left(FieldId,1) = "," Then FieldId = Right(FieldId,Len(FieldId)-1)
	
		Response.Write FieldStr & "<input type=""hidden"" name=""FieldId"& rsClass("id") &""" value="""& FieldId &""" />"
	
		Response.Write "</table></div>"
		
	rsClass.MoveNext         
	Loop 
	
	'/�Զ����ֶ� ===================================================================================
%>
		<script type="text/javascript">
			ShowField("<%=ClassID%>");
		</script>		</td></tr>
        <tr>
          <td width="82" height="30" align="center"></td>
          <td colspan="2" valign="bottom" style="padding:10px 0;"><input name="Submit" type="submit" class="btn2" value=" �� �� " ></td>
        </tr>
      </table>
</form>
<%
	CloseConn
%>
</body>
</html>