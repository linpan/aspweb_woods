<!--#include file="../../Include/Class_Conn.asp"-->
<!--#include file="../../Include/Class_Main.asp"-->
<!--#include file="../p8_Check.asp"-->
<%
	If Request("Submit")<>"" Then
		Dim i,ClassID,BigClass,SmallClass,Title,Source,AddDate,TitleColor,Url,KeyWord,Pic,SmallPic,Content,sFieldId,FieldContent
		ClassID    = Trim(Request.Form("ClassID"))
		Title      = Trim(Request.Form("Title"))
		Source     = Trim(Request.Form("Source"))
		AddDate    = Trim(Request.Form("AddDate"))
		TitleColor = Trim(Request.Form("TitleColor"))
		Url        = Trim(Request.Form("Url"))
		KeyWord    = Trim(Request.Form("KeyWord"))
		Pic        = Trim(Request.Form("Pic"))
		SmallPic   = Trim(Request.Form("SmallPic"))
		Content    = Trim(Request.Form("Content"))
		
		'�����Զ����ֶ� =================================================================================
		FieldId    = Trim(Request.Form("FieldId"& ClassID &""))
		sFieldId   = Split(FieldId,",")'����Ӧ�ý�����Щ�ֶ���
		
		For i = 0 To Ubound(sFieldId)
			TagName    = Replace(sFieldId(i),"Fz_FieldId"& ClassID & "_" ,"") '���ʱ�����ʶ��
			FieldContent = FieldContent & "{$"& TagName &"$}"& Request.Form(sFieldId(i)) &"{$/"& TagName &"$}"
		Next
		'/�����Զ����ֶ� =================================================================================

		If ClassID="" Then
			Response.Write "<script>alert(""��ѡ���࣡"");window.history.back();</script>"
			Response.End()
		End If

		If Title="" Then
			Response.Write "<script>alert(""����д���⣡"");window.history.back();</script>"
			Response.End()
		End If
		
		KeyWord = Replace(KeyWord," ",",")
		KeyWord = Replace(KeyWord,"|",",")
		KeyWord = Replace(KeyWord,"��",",")
		KeyWord = Replace(KeyWord,";",",")
		While Instr(KeyWord,",,")
			KeyWord = Replace(KeyWord,",,",",")
		Wend
		If Left(KeyWord,1)  = "," Then KeyWord = Right(KeyWord,Len(KeyWord)-1)
		If Right(KeyWord,1) = "," Then KeyWord = Left(KeyWord,Len(KeyWord)-1)
		
		'ȡ�ô����С��ID
		Set rs = Server.Createobject("Adodb.RecordSet")
		rs.open "Select ClassLevel,ParentID From p8_Class Where id="& ClassID &"",Conn,1,1
		
			If Not rs.Eof Then
				If rs("ClassLevel") = 1 Then '���ѡ��ķ�����һ������ֱ���������´���
					BigClass = ClassID
				End If
				
				If rs("ClassLevel") = 2 Then '���ѡ��ķ����Ƕ����������һ������ID
					BigClass   = rs("ParentID")
					SmallClass = ClassID
				End If
			End If

		rs.Close
		Set rs = Nothing
		
		'��¼ʹ��ϰ��
		Call History("������Դ",Source,"")

		Set rs=Server.CreateObject("Adodb.Recordset")
		rs.Open "Select * From p8_News",Conn,1,3
		rs.AddNew
		
		rs("BigClass")     = BigClass
		rs("SmallClass")   = SmallClass
		rs("Title")        = Title
		rs("Source")       = Source
		rs("AddDate")      = AddDate
		rs("TitleColor")   = TitleColor
		rs("Url")          = Url
		rs("Pic")          = Pic
		rs("SmallPic")     = SmallPic
		rs("Content")      = Content
		rs("KeyWord")      = KeyWord
		rs("Admin")        = Request.Cookies("Admin")("s_Name")
		rs("FieldContent") = FieldContent

		rs.Update
		rs.Close
		Set rs=Nothing
		CloseConn
		Response.Redirect "News_Add.asp?ClassID="& ClassID &"&Tip=��ӳɹ���"
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
<title>�������</title>
<script type="text/javascript">top.window.aTitle.innerText='�������'</script>
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
	if (document.getElementById("Title").value=="") {
        alert("����д����");
		document.getElementById("Title").focus();
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
		rs.open "Select id From p8_Class Where ClassType='����' Order By id Desc",Conn,1,1
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
	<form id="AddFrom" name="AddFrom" method="post" action="News_Add.asp" onSubmit="return CheckForm();">
      <table width="100%" border="0" cellpadding="0" cellspacing="0" style="margin:5px 0;">
        <tr>
          <td width="82" height="30" align="right">���ࣺ<span class="cRed">*</span>&nbsp;</td>
          <td>
		<select name="ClassID" id="ClassID" style="width:120px;" class="ipt5" onChange="ShowField(this.value)">
		 <option value=""></option>
		  <%
			Set rs = Server.Createobject("Adodb.RecordSet")
			rs.open "Select id,ClassName From p8_Class Where ClassType='����' And ClassLevel=1 Order By id Desc",Conn,1,1
			
				Do While Not rs.Eof 
					
					If Clng(ClassID) = Clng(rs("id")) Then
						Response.Write "<option value="""& rs("id") &""" selected=""selected"">"& rs("ClassName") &"</option>"
					Else
						Response.Write "<option value="""& rs("id") &""">"& rs("ClassName") &"</option>"
					End If
					
					Set rsnxt = Server.CreateObject("ADODB.RecordSet")
					rsnxt.open "Select id,ClassName From p8_Class Where ClassType='����' And ClassLevel=2 And ParentID="& rs("id") &" Order By id Desc",Conn,1,1
					
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
          <td align="right">��Դ��</td>
          <td width="115">
		  <div id="SourceMenu" style="position:relative; display:none;">
		  		<div style="z-index:999; width:100px; position:absolute; left:0; top:19px; background-color:#fff; border:1px solid #8db2e3; padding-top:3px; overflow:hidden;">
					<table width="110" border="0" cellspacing="0" cellpadding="0">
					<%
						Set rs = Server.CreateObject("Adodb.Recordset")
						rs.Open "Select Top 20 His_Name From p8_History Where His_Class = '������Դ' And His_User = '"& Request.Cookies("Admin")("s_User") &"' Order By His_Hit Desc",Conn,1,1
						
						n = 1
						Do While Not rs.Eof
							If n<=5 Then
								SourceColor = "color:#ff5400;"
							Else
								SourceColor = ""
							End If
					%>
						<tr><td height="20" onMouseDown="document.getElementById('Source').value='<%=rs("His_Name")%>'" onMouseOver="this.style.backgroundColor='#e7ffa6'" onMouseOut="this.style.backgroundColor='#fff'" style="cursor:pointer;<%=SourceColor%>">&nbsp;<%=rs("His_Name")%></td></tr>
					<%
							n = n + 1
							rs.MoveNext
						Loop
						
						rs.Close
						Set rs = Nothing
					%>
					</table>
				</div>
		  </div>
		  <input name="Source" onFocus="document.getElementById('SourceMenu').style.display='block';document.getElementById('TitleColor').style.display='none';" onBlur="document.getElementById('SourceMenu').style.display='none';document.getElementById('TitleColor').style.display='block';" type="text" class="ipt4" id="Source" style="width:100px; background:url(../images/Jian_1.gif) 92px 11px no-repeat;" autocomplete="off"></td>
          <td width="50" align="right">���ڣ�</td>
          <td><input name="AddDate" type="text" class="ipt4" id="AddDate" onFocus="setday(this)" value="<%=Now()%>" maxlength="50" style="cursor:pointer;width:140px;"></td>
          <td rowspan="3">
		  <div id="ShowPic" style="display:;"><img id="ShowImg" onerror="this.src='../images/NoPhoto.gif'" src="../images/NoPhoto.gif" style="height:75px; border:1px solid #E8E8E8; padding:2px; margin:5px;" /></div>		  </td>
        </tr>
        <tr>
          <td width="82" height="30" align="right">���⣺<span class="cRed">*</span>&nbsp;</td>
          <td width="302">
			<input name="Title" type="text" class="ipt4" id="Title" style="width:290px;" maxlength="100"></td>
          <td width="70" align="right">��ɫ��</td>
          <td>
		  <select name="TitleColor" id="TitleColor" style="width:100px;" class="ipt5">
            <option value=""></option>
            <option value="#FF0000" style="background-color:#FF0000; color:#fff;">��ɫ</option>
			<option value="#3333FF" style="background-color:#3333FF; color:#fff;">��ɫ</option>
			<option value="#009900" style="background-color:#009900; color:#fff;">��ɫ</option>
			<option value="#FF6600" style="background-color:#FF6600; color:#fff;">��ɫ</option>
			<option value="#9933CC" style="background-color:#9933CC; color:#fff;">��ɫ</option>
          </select></td>
          <td align="right">���ӣ�</td>
          <td width="302"><input name="Url" type="text" class="ipt4" id="Url" style="width:300px;" maxlength="100"></td>
        </tr>
        <tr>
          <td height="30" align="right">�ؼ��֣�&nbsp;&nbsp;</td>
          <td colspan="3"><input name="KeyWord" type="text" class="ipt4" id="KeyWord" style="width:290px;" maxlength="100">
              <span class="cGray">�����ʹ��,����</span></td>
          <td align="right">ͼƬ��</td>
          <td width="302"><input name="Pic" type="text" class="ipt4" id="Pic" style="width:300px; display:none;"  readonly="readonly" onChange="doChange(this,SmallPic)">
              <select id="SmallPic" name="SmallPic" size="1" class="ipt4" style="width:300px;" onChange="document.getElementById('ShowPic').style.display='block';document.getElementById('ShowImg').src=this.value;">
              </select>          </td>
        </tr>
        <tr>
          <td width="82" height="30" align="right">���ݣ�<span class="cRed">*</span>&nbsp;</td>
          <td colspan="6" style="padding-right:20px;"><textarea name="Content" style="display:none"></textarea>
              <IFRAME ID="eWebEditor1" src="../Include/Editer/ewebeditor.htm?id=Content&style=coolblue&savepathfilename=Pic" frameborder="0" scrolling="no" width="100%" height="430"></IFRAME></td>
        </tr>
        <tr>
          <td colspan="7" align="right">
<%
	'�Զ����ֶ� ====================================================================================
	Dim ClassNum,FieldStr,FieldId,m

	Set rsClass = Server.Createobject("Adodb.RecordSet")
	rsClass.open "Select id From p8_Class Where ClassType='����' Order By id Desc",Conn,1,1 '��ȡ���з���ID,���ݷ���ID�г����з����Զ����ֶ�

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
		</script>
		</td></tr>
        <tr>
          <td width="82" height="30" align="center"></td>
          <td colspan="6" valign="bottom" style="padding:20px 0;"><input name="Submit" type="submit" class="btn2" value=" �� �� " ></td>
        </tr>
      </table>
</form>
<%
	CloseConn
%>
</body>
</html>