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
		
		'接收自定义字段 =================================================================================
		FieldId    = Trim(Request.Form("FieldId"& ClassID &""))
		sFieldId   = Split(FieldId,",") '告诉应该接收哪些字段名
		
		For i = 0 To Ubound(sFieldId)
			TagName    = Replace(sFieldId(i),"Fz_FieldId"& ClassID & "_" ,"") '入库时清除标识符
			FieldContent = FieldContent & "{$"& TagName &"$}"& Request.Form(sFieldId(i)) &"{$/"& TagName &"$}"
		Next
		'/接收自定义字段 =================================================================================
		
		id        = Request("id")
		Page      = Request("Page")
		s_ClassID = Request("s_ClassID")
		s_Title   = Request("s_Title")
		s_Source  = Request("s_Source")
		s_AddDate = Request("s_AddDate")
		s_News_px = Request("s_News_px")

		If ClassID="" Then
			Response.Write "<script>alert(""请选分类！"");window.history.back();</script>"
			Response.End()
		End If

		If Title="" Then
			Response.Write "<script>alert(""请填写标题！"");window.history.back();</script>"
			Response.End()
		End If
		
		KeyWord = Replace(KeyWord," ",",")
		KeyWord = Replace(KeyWord,"|",",")
		KeyWord = Replace(KeyWord,"，",",")
		KeyWord = Replace(KeyWord,";",",")
		While Instr(KeyWord,",,")
			KeyWord = Replace(KeyWord,",,",",")
		Wend
		If Left(KeyWord,1)  = "," Then KeyWord = Right(KeyWord,Len(KeyWord)-1)
		If Right(KeyWord,1) = "," Then KeyWord = Left(KeyWord,Len(KeyWord)-1)
		
		'取得大类和小类ID
		Set rs = Server.Createobject("Adodb.RecordSet")
		rs.open "Select ClassLevel,ParentID From p8_Class Where id="& ClassID &"",Conn,1,1
		
			If Not rs.Eof Then
				If rs("ClassLevel") = 1 Then '如果选择的分类是一级，则直接设置文章大类
					BigClass = ClassID
				End If
				
				If rs("ClassLevel") = 2 Then '如果选择的分类是二级，则查找一级分类ID
					BigClass   = rs("ParentID")
					SmallClass = ClassID
				End If
			End If

		rs.Close
		Set rs = Nothing
		
		'记录使用习惯
		Call History("文章来源",Source,"")

		Set rs=Server.CreateObject("Adodb.Recordset")
		rs.Open "Select * From p8_News Where id = " & id,Conn,1,3
		
		rs("BigClass")   = BigClass
		rs("SmallClass") = SmallClass
		rs("Title")      = Title
		rs("Source")     = Source
		rs("AddDate")    = AddDate
		rs("TitleColor") = TitleColor
		rs("Url")        = Url
		rs("Pic")        = Pic
		rs("SmallPic")   = SmallPic
		rs("Content")    = Content
		rs("KeyWord")    = KeyWord
		rs("Admin")      = Request.Cookies("Admin")("s_Name")
		rs("FieldContent") = FieldContent
		
		rs.Update
		rs.Close
		Set rs=Nothing
		CloseConn
		Response.Redirect "News_List.asp?Tip=修改成功！&Page="& Page &"&ClassID="& s_ClassID &"&Title="& s_Title &"&Source="& s_Source &"&AddDate="& s_AddDate &"&News_px="& s_News_px
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
<title>修改文章</title>
<script type="text/javascript">top.window.aTitle.innerText='修改文章'</script>
<link href="../css/Public.css" rel="stylesheet" type="text/css" /> 
<script type="text/javascript" src="../Include/TipBox.js"></script>
<script type="text/javascript" src="../Include/calendar.js"></script>
<script type="text/javascript" src="../Include/Pub.js"></script>
<script type="text/javascript">
function CheckForm(){
    if (document.getElementById("ClassID").value=="") {
        alert("请选择分类");
		document.getElementById("ClassID").focus();
		return false;
    }
	if (document.getElementById("Title").value=="") {
        alert("请填写标题");
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
		rs.open "Select id From p8_Class Where ClassType='文章' Order By id Desc",Conn,1,1
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
	
	Dim Page,s_ClassID,s_Title,s_Source,s_AddDate,s_News_px
	id        = Request("id")
	Page      = Request("Page")
	s_ClassID = Request("ClassID")
	s_Title   = Request("Title")
	s_Source  = Request("Source")
	s_AddDate = Request("AddDate")
	s_News_px = Request("News_px")
	
	Set rs = Server.CreateObject ("ADODB.Recordset")
	rs.Open "Select * From p8_News Where id="& id &"",Conn,1,1
%>
	<form id="AddFrom" name="AddFrom" method="post" action="News_Edit.asp" onSubmit="return CheckForm();">
	<input name="id" type="hidden" value="<%=id%>">
	<input name="Page" type="hidden" value="<%=Page%>">
	<input name="s_ClassID" type="hidden" value="<%=s_ClassID%>">
	<input name="s_Title" type="hidden" value="<%=s_Title%>">
	<input name="s_Source" type="hidden" value="<%=s_Source%>">
	<input name="s_AddDate" type="hidden" value="<%=s_AddDate%>">
	<input name="s_News_px" type="hidden" value="<%=s_News_px%>">
      <table width="100%" border="0" cellpadding="0" cellspacing="0" style="margin:5px 0;">
        <tr>
          <td width="82" height="30" align="right">分类：<span class="cRed">*</span>&nbsp;</td>
          <td>
		<select name="ClassID" id="ClassID" style="width:120px;" class="ipt5" onChange="ShowField(this.value)">
		 <option value=""></option>
		  <%
			Set rsc = Server.Createobject("Adodb.RecordSet")
			rsc.open "Select id,ClassName From p8_Class Where ClassType='文章' And ClassLevel=1 Order By id Desc",Conn,1,1
			
				Do While Not rsc.Eof
					
					If Clng(rs("BigClass")) = Clng(rsc("id")) Then
						Response.Write "<option value="""& rsc("id") &""" selected=""selected"">"& rsc("ClassName") &"</option>"
					Else
						Response.Write "<option value="""& rsc("id") &""">"& rsc("ClassName") &"</option>"
					End If
					
					Set rsnxt = Server.CreateObject("ADODB.RecordSet")
					rsnxt.open "Select id,ClassName From p8_Class Where ClassType='文章' And ClassLevel=2 And ParentID="& rsc("id") &" Order By id Desc",Conn,1,1
					
					Do While Not rsnxt.Eof 

						'If rs("SmallClass")<>"" Then
							If (rs("SmallClass")) = Clng(rsnxt("id")) Then
								Response.Write "<option value="""& rsnxt("id") &""" selected=""selected"">├ "& rsnxt("ClassName") &"</option>"
							Else
								Response.Write "<option value="""& rsnxt("id") &""">├ "& rsnxt("ClassName") &"</option>"
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
		</select></td>
          <td align="right">来源：</td>
          <td width="115">
		  <div id="SourceMenu" style="position:relative; display:none;">
		  		<div style="z-index:999; width:100px; position:absolute; left:0; top:19px; background-color:#fff; border:1px solid #8db2e3; padding-top:3px; overflow:hidden;">
					<table width="110" border="0" cellspacing="0" cellpadding="0">
					<%
						Set rsc = Server.CreateObject("Adodb.Recordset")
						rsc.Open "Select Top 20 His_Name From p8_History Where His_Class = '文章来源' And His_User = '"& Request.Cookies("Admin")("s_User") &"' Order By His_Hit Desc",Conn,1,1
						
						n = 1
						Do While Not rsc.Eof
							If n<=5 Then
								SourceColor = "color:#ff5400;"
							Else
								SourceColor = ""
							End If
					%>
						<tr><td height="20" onMouseDown="document.getElementById('Source').value='<%=rsc("His_Name")%>'" onMouseOver="this.style.backgroundColor='#e7ffa6'" onMouseOut="this.style.backgroundColor='#fff'" style="cursor:pointer;<%=SourceColor%>">&nbsp;<%=rsc("His_Name")%></td></tr>
					<%
							n = n + 1
							rsc.MoveNext
						Loop
						
						rsc.Close
						Set rsc = Nothing
					%>
					</table>
				</div>
		  </div>
		  <input name="Source" value="<%=rs("Source")%>" onFocus="document.getElementById('SourceMenu').style.display='block';document.getElementById('TitleColor').style.display='none';" onBlur="document.getElementById('SourceMenu').style.display='none';document.getElementById('TitleColor').style.display='block';" type="text" class="ipt4" id="Source" style="width:100px; background:url(../images/Jian_1.gif) 92px 11px no-repeat;" autocomplete="off"></td>
          <td width="50" align="right">日期：</td>
          <td><input name="AddDate" value="<%=rs("AddDate")%>" type="text" class="ipt4" id="AddDate" onFocus="setday(this)" maxlength="50" style="cursor:pointer;width:140px;"></td>
          <td rowspan="3">
		  <div id="ShowPic" style="display:;"><img id="ShowImg" onerror="this.src='../images/NoPhoto.gif'" src="../images/NoPhoto.gif" style="height:75px; border:1px solid #E8E8E8; padding:2px; margin:5px;" /></div>		  </td>
        </tr>
        <tr>
          <td width="82" height="30" align="right">标题：<span class="cRed">*</span>&nbsp;</td>
          <td width="302">
			<input name="Title" value="<%=rs("Title")%>" type="text" class="ipt4" id="Title" style="width:290px;" maxlength="100"></td>
          <td width="70" align="right">颜色：</td>
          <td>
		  <select name="TitleColor" id="TitleColor" style="width:100px;" class="ipt5">
            <option value=""></option>
            <option value="#FF0000" style="background-color:#FF0000; color:#fff;" <%If rs("TitleColor") = "#FF0000" Then Response.Write "selected=""selected"""%>>红色</option>
			<option value="#3333FF" style="background-color:#3333FF; color:#fff;" <%If rs("TitleColor") = "#3333FF" Then Response.Write "selected=""selected"""%>>蓝色</option>
			<option value="#009900" style="background-color:#009900; color:#fff;" <%If rs("TitleColor") = "#009900" Then Response.Write "selected=""selected"""%>>绿色</option>
			<option value="#FF6600" style="background-color:#FF6600; color:#fff;" <%If rs("TitleColor") = "#FF6600" Then Response.Write "selected=""selected"""%>>黄色</option>
			<option value="#9933CC" style="background-color:#9933CC; color:#fff;" <%If rs("TitleColor") = "#9933CC" Then Response.Write "selected=""selected"""%>>紫色</option>
          </select></td>
          <td align="right">链接：</td>
          <td width="302"><input name="Url" value="<%=rs("Url")%>" type="text" class="ipt4" id="Url" style="width:300px;" maxlength="100"></td>
        </tr>
        <tr>
          <td height="30" align="right">关键字：&nbsp;&nbsp;</td>
          <td colspan="3"><input name="KeyWord" value="<%=rs("KeyWord")%>" type="text" class="ipt4" id="KeyWord" style="width:290px;" maxlength="100">
              <span class="cGray">多个请使用,隔开</span></td>
          <td align="right">图片：</td>
          <td width="302"><input name="Pic" type="text" class="ipt4" id="Pic" style="width:300px; display:none;"  readonly="readonly" onChange="doChange(this,SmallPic)" value="<%=rs("Pic")%>">
              <select id="SmallPic" name="SmallPic" size="1" class="ipt4" style="width:300px;" onChange="document.getElementById('ShowPic').style.display='block';document.getElementById('ShowImg').src=this.value;">
              </select>
			<script type="text/javascript">
				  doChange(document.getElementById('Pic'),document.getElementById('SmallPic'),"<%=rs("SmallPic")%>");
				  document.getElementById('ShowPic').style.display='block';
				  document.getElementById('ShowImg').src='<%=rs("SmallPic")%>';
			  </script>	
			  </td>
        </tr>
        <tr>
          <td width="82" height="30" align="right">内容：<span class="cRed">*</span>&nbsp;</td>
          <td colspan="6" style="padding-right:20px;"><textarea name="Content" style="display:none"><%=rs("Content")%></textarea>
              <IFRAME ID="eWebEditor1" src="../Include/Editer/ewebeditor.htm?id=Content&style=coolblue&savepathfilename=Pic" frameborder="0" scrolling="no" width="100%" height="430"></IFRAME></td>
        </tr>
        <tr>
          <td colspan="7" align="right">
<%
	'自定义字段 ====================================================================================
	Dim ClassNum,FieldStr,FieldId,m

	Set rsClass = Server.Createobject("Adodb.RecordSet")
	rsClass.open "Select id From p8_Class Where ClassType='文章' Order By id Desc",Conn,1,1

	Do While Not rsClass.Eof '读取所有分类ID,根据分类ID列出所有分类自定义字段
	
		Response.Write "<div id=""FieldDiv"& rsClass("id") &""" style=""display:none;""><table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		
		Set rsNum = Server.Createobject("Adodb.RecordSet") '获取标识码
		rsNum.open "Select Num From p8_Class Where id="& rsClass("id"),Conn,1,1
		ClassNum = rsNum("Num")
		rsNum.Close
		Set rsNum = Nothing
		
		Set rsField = Server.Createobject("Adodb.RecordSet") '读取该分类下的所有自定义字段
		rsField.open "Select id,FieldName,Variable,FieldType,MaxLen,Width,Height,Content,Options From p8_Field Where ClassNum='"& ClassNum &"' Order By id Desc",Conn,1,1
		
			FieldStr = ""
			FieldId  = ""
			Fz_FieldId = "Fz_FieldId"& rsClass("id") &"_" '生成应该接受的字段，并加上分类ID进行指定接受，放置不同分类同名字段

			Do While Not rsField.Eof 
				FieldId    = FieldId & "," & Fz_FieldId & rsField("Variable")
				FieldValue = GetField(rsField("Variable"),id)
				
				If rsField("MaxLen") <> 0 Then 
					MaxLen = " maxlength="""& rsField("MaxLen") &""""
				End If
				
				'单行文本框 =============================================
				If rsField("FieldType") = "text" Then
					FieldStr = FieldStr & "<tr><td width=""82"" height=""30"" align=""right"">"& rsField("FieldName") &"：&nbsp;</td><td colspan=""6""><input name="""& Fz_FieldId & rsField("Variable") &""" type=""text"" class=""ipt4"" value="""& FieldValue &""" style=""width:"& rsField("Width") &";"" "& MaxLen &"></td></tr>"
				End If
				
				'多行文本框 =============================================
				If rsField("FieldType") = "textarea" Then
					FieldStr = FieldStr & "<tr><td width=""82"" height=""30"" align=""right"">"& rsField("FieldName") &"：&nbsp;</td><td colspan=""6"" style=""padding-top:5px;""><textarea name="""& Fz_FieldId & rsField("Variable") &""" class=""ipt3"" style=""width:"& rsField("Width") &"px; height:"& rsField("Height") &"px;"">"& FieldValue &"</textarea></td></tr>"
				End If
				
				'单选框 =================================================
				If rsField("FieldType") = "radio" Then
					
					FieldStr = FieldStr & "<tr><td width=""82"" height=""30"" align=""right"">"& rsField("FieldName") &"：&nbsp;</td><td colspan=""6"">"
					
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
				
				'复选框 =================================================
				If rsField("FieldType") = "checkbox" Then
					FieldStr = FieldStr & "<tr><td width=""82"" height=""30"" align=""right"">"& rsField("FieldName") &"：&nbsp;</td><td colspan=""6"">"
					
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
				
				'下拉框 =================================================
				If rsField("FieldType") = "select" Then
					FieldStr = FieldStr & "<tr><td width=""82"" height=""30"" align=""right"">"& rsField("FieldName") &"：&nbsp;</td><td colspan=""6"">"
					
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
	
	'/自定义字段 ===================================================================================
	
	If rs("SmallClass") <> "" Then
		ClassID = rs("SmallClass")
	Else
		ClassID = rs("BigClass")
	End If
%>
		<script type="text/javascript">
			ShowField("<%=ClassID%>");
		</script>
		</td></tr>
        <tr>
          <td width="82" height="30" align="center">&nbsp;</td>
          <td colspan="6" valign="bottom" style="padding:20px 0;"><input name="Submit" type="submit" class="btn2" value=" 修 改 " ></td>
        </tr>
      </table>
</form>
<%
	CloseConn
%>
</body>
</html>