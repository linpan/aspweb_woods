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
<title>�޸�ͼƬ</title>
<script type="text/javascript">top.window.aTitle.innerText='�޸�ͼƬ'</script>
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
	
	'-----------------------------------------------------------------------
	
	Dim Page,s_ClassID,s_PicName,s_Content,s_Pic_px
	id        = Request("id")
	Page      = Request("Page")
	s_ClassID = Request("ClassID")
	s_PicName = Request("PicName")
	s_Content = Request("Content")
	s_Pic_px  = Request("Pic_px")
	
	Set rs = Server.CreateObject ("ADODB.Recordset")
	rs.Open "Select * From p8_Pic Where id="& Clng(id) &"",Conn,1,1
%>
	<form id="AddFrom" name="AddFrom" enctype="multipart/form-data" method="post" action="Pic_Edit_Sql.asp" onSubmit="return CheckForm();">
	<input name="id" type="hidden" value="<%=id%>">
	<input name="Page" type="hidden" value="<%=Page%>">
	<input name="s_ClassID" type="hidden" value="<%=s_ClassID%>">
	<input name="s_PicName" type="hidden" value="<%=s_PicName%>">
	<input name="s_Content" type="hidden" value="<%=s_Content%>">
	<input name="s_Pic_px" type="hidden" value="<%=s_Pic_px%>">
      <table width="100%" border="0" cellpadding="0" cellspacing="0" style="margin:5px 0;">
        <tr>
          <td width="82" height="30" align="right">���ࣺ<span class="cRed">*</span>&nbsp;</td>
          <td width="317">
		<select name="ClassID" id="ClassID" style="width:120px;" class="ipt5" onChange="ShowField(this.value)">
		 <option value=""></option>
		  <%
			Set rsc = Server.Createobject("Adodb.RecordSet")
			rsc.open "Select id,ClassName From p8_Class Where ClassType='ͼƬ' And ClassLevel=1 Order By id Desc",Conn,1,1
			
				Do While Not rsc.Eof
					
					If Clng(rs("BigClass")) = Clng(rsc("id")) Then
						Response.Write "<option value="""& rsc("id") &""" selected=""selected"">"& rsc("ClassName") &"</option>"
					Else
						Response.Write "<option value="""& rsc("id") &""">"& rsc("ClassName") &"</option>"
					End If
					
					Set rsnxt = Server.CreateObject("ADODB.RecordSet")
					rsnxt.open "Select id,ClassName From p8_Class Where ClassType='ͼƬ' And ClassLevel=2 And ParentID="& rsc("id") &" Order By id Desc",Conn,1,1
					
					Do While Not rsnxt.Eof 

						'If rs("SmallClass")<>"" Then
							If (rs("SmallClass")) = Clng(rsnxt("id")) Then
								Response.Write "<option value="""& rsnxt("id") &""" selected=""selected"">�� "& rsnxt("ClassName") &"</option>"
							Else
								Response.Write "<option value="""& rsnxt("id") &""">�� "& rsnxt("ClassName") &"</option>"
							End If
						'End If
					
						rsnxt.MoveNext         
					Loop 
					
					rsnxt.Close
					Set rsnxt = Nothing

				
				rsc.MoveNext         
				Loop 

			rsc.Close
			Set rsc = Nothing
		  %>
		</select>		</td>
          <td width="769" rowspan="3"><img id="ShowImg" onerror="this.src='../images/NoPhoto.gif'" src="<%=rs("SmallPic")%>" style="height:75px; border:1px solid #E8E8E8; padding:2px; margin:5px;" /></td>
        </tr>
        <tr>
          <td width="82" height="30" align="right">���ƣ�<span class="cRed">*</span>&nbsp;</td>
          <td>
			<input name="PicName" type="text" class="ipt4" id="PicName" style="width:290px;" maxlength="100" value="<%=rs("PicName")%>"></td>
        </tr>
        <tr>
          <td width="82" height="30" align="right">����ͼ��<span class="cRed">*</span>&nbsp;</td>
          <td>
		  <input name="Pic" type="text" class="ipt4" id="Pic" style="width:300px; display:none;"  readonly="readonly" value="<%=Pic%>">
		  <input name="SmallPic" type="file" class="ipt4" id="SmallPic" style="width:290px;"  value="" size="20" maxlength="50" onChange="$Get('ShowImg').src=this.value" /></td>
        </tr>
        <tr>
          <td width="82" height="30" align="right">���ݣ�<span class="cRed">*</span>&nbsp;</td>
          <td colspan="2" style="padding-right:20px;"><textarea name="Content" style="display:none"><%=rs("Content")%></textarea>
              <IFRAME ID="eWebEditor1" src="../Include/Editer/ewebeditor.htm?id=Content&style=coolblue&savepathfilename=Pic" frameborder="0" scrolling="no" width="100%" height="430"></IFRAME></td>
        </tr>
        <tr>
          <td colspan="3" align="right">
<%
	'�Զ����ֶ� ====================================================================================
	Dim ClassNum,FieldStr,FieldId,m

	Set rsClass = Server.Createobject("Adodb.RecordSet")
	rsClass.open "Select id From p8_Class Where ClassType='ͼƬ' Order By id Desc",Conn,1,1

	Do While Not rsClass.Eof '��ȡ���з���ID,���ݷ���ID�г����з����Զ����ֶ�
	
		Response.Write "<div id=""FieldDiv"& rsClass("id") &""" style=""display:none;""><table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		
		Set rsNum = Server.Createobject("Adodb.RecordSet") '��ȡ��ʶ��
		rsNum.open "Select Num From p8_Class Where id="& rsClass("id"),Conn,1,1
		ClassNum = rsNum("Num")
		rsNum.Close
		Set rsNum = Nothing
		
		Set rsField = Server.Createobject("Adodb.RecordSet") '��ȡ�÷����µ������Զ����ֶ�
		rsField.open "Select id,FieldName,Variable,FieldType,MaxLen,Width,Height,Content,Options From p8_Field Where ClassNum='"& ClassNum &"' Order By id Desc",Conn,1,1
		
			FieldStr = ""
			FieldId  = ""
			Fz_FieldId = "Fz_FieldId"& rsClass("id") &"_" '����Ӧ�ý��ܵ��ֶΣ������Ϸ���ID����ָ�����ܣ����ò�ͬ����ͬ���ֶ�

			Do While Not rsField.Eof 
				FieldId    = FieldId & "," & Fz_FieldId & rsField("Variable")
				FieldValue = GetPicField(rsField("Variable"),id)
				
				If rsField("MaxLen") <> 0 Then 
					MaxLen = " maxlength="""& rsField("MaxLen") &""""
				End If
				
				'�����ı��� =============================================
				If rsField("FieldType") = "text" Then
					FieldStr = FieldStr & "<tr><td width=""82"" height=""30"" align=""right"">"& rsField("FieldName") &"��&nbsp;</td><td colspan=""6""><input name="""& Fz_FieldId & rsField("Variable") &""" type=""text"" class=""ipt4"" value="""& FieldValue &""" style=""width:"& rsField("Width") &";"" "& MaxLen &"></td></tr>"
				End If
				
				'�����ı��� =============================================
				If rsField("FieldType") = "textarea" Then
					FieldStr = FieldStr & "<tr><td width=""82"" height=""30"" align=""right"">"& rsField("FieldName") &"��&nbsp;</td><td colspan=""6"" style=""padding-top:5px;""><textarea name="""& Fz_FieldId & rsField("Variable") &""" class=""ipt3"" style=""width:"& rsField("Width") &"px; height:"& rsField("Height") &"px;"">"& FieldValue &"</textarea></td></tr>"
				End If
				
				'��ѡ�� =================================================
				If rsField("FieldType") = "radio" Then
					
					FieldStr = FieldStr & "<tr><td width=""82"" height=""30"" align=""right"">"& rsField("FieldName") &"��&nbsp;</td><td colspan=""6"">"
					
					Options = Split(rsField("Options"),chr(13))
					
					For j = 0 To Ubound(Options)
						If Instr(FieldValue,"|"& Replace(Trim(Options(j)),chr(10),"") &"|") Then
							FieldStr = FieldStr & "<input type=""radio"" name="""& Fz_FieldId & rsField("Variable") &""" value=""|"& Replace(Trim(Options(j)),chr(10),"") &"|"" checked=""checked"" />"& Trim(Options(j)) &" "
						Else
							FieldStr = FieldStr & "<input type=""radio"" name="""& Fz_FieldId & rsField("Variable") &""" value=""|"& Replace(Trim(Options(j)),chr(10),"") &"|"" />"& Trim(Options(j)) &" "
						End If
					Next
					
					FieldStr = FieldStr & "</td></tr>"
				End If
				
				'��ѡ�� =================================================
				If rsField("FieldType") = "checkbox" Then
					FieldStr = FieldStr & "<tr><td width=""82"" height=""30"" align=""right"">"& rsField("FieldName") &"��&nbsp;</td><td colspan=""6"">"
					
					Options = Split(rsField("Options"),chr(13))

					For j = 0 To Ubound(Options)
						If Instr(FieldValue,"|"& Replace(Trim(Options(j)),chr(10),"") &"|") Then
							FieldStr = FieldStr & "<input type=""checkbox"" name="""& Fz_FieldId & rsField("Variable") &""" value=""|"& Replace(Trim(Options(j)),chr(10),"") &"|"" checked=""checked"" />"& Trim(Options(j)) &" " 
						Else
							FieldStr = FieldStr & "<input type=""checkbox"" name="""& Fz_FieldId & rsField("Variable") &""" value=""|"& Replace(Trim(Options(j)),chr(10),"") &"|"" />"& Trim(Options(j)) &" " 
						End If
					Next
					
					FieldStr = FieldStr & "</td></tr>"
				End If
				
				'������ =================================================
				If rsField("FieldType") = "select" Then
					FieldStr = FieldStr & "<tr><td width=""82"" height=""30"" align=""right"">"& rsField("FieldName") &"��&nbsp;</td><td colspan=""6"">"
					
					Options = Split(rsField("Options"),chr(13))
					
					FieldStr = FieldStr & "<select style=""width:"& rsField("Width") &";"" name="""& Fz_FieldId & rsField("Variable") &"""><option value=""""></option>"
					
					For j = 0 To Ubound(Options)
						If Instr(FieldValue,"|"& Replace(Trim(Options(j)),chr(10),"") &"|") Then
							FieldStr = FieldStr & "<option value=""|"& Replace(Trim(Options(j)),chr(10),"") &"|"" selected=""selected"">"& Trim(Options(j)) &"</option>" 
						Else
							FieldStr = FieldStr & "<option value=""|"& Replace(Trim(Options(j)),chr(10),"") &"|"">"& Trim(Options(j)) &"</option>" 
						End If
					Next
					
					FieldStr = FieldStr & "</select>"
					
					FieldStr = FieldStr & "</td></tr>"
				End If
			
			rsField.MoveNext         
			Loop 
	
		rsField.Close
		Set rsField = Nothing
		
		If Left(FieldId,1) = "," Then FieldId = Right(FieldId,Len(FieldId)-1)
	
		Response.Write FieldStr & "<input type=""hidden"" name=""FieldId"& rsClass("id") &""" value="""& FieldId &""" />"
	
		Response.Write "</table></div>"
		
	rsClass.MoveNext         
	Loop 
	
	'/�Զ����ֶ� ===================================================================================
	
	If rs("SmallClass") <> "" Then
		ClassID = rs("SmallClass")
	Else
		ClassID = rs("BigClass")
	End If
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