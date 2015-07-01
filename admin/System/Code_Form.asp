<!--#include file="../../Include/Class_Conn.asp"-->
<!--#include file="../../Include/Class_Main.asp"-->
<% Super=1 '只允许超级管理员访问该页 %>
<!--#include file="../p8_Check.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>表单 - 数据调用</title>
<script type="text/javascript">top.window.aTitle.innerText='数据调用'</script>
<meta http-equiv="pragma" content="no-cache" />
<meta http-equiv="Cache-Control" content="no-cache, must-revalidate" />
<meta http-equiv="expires" content="Wed, 26 Feb 1997 08:21:57 GMT" />
<meta http-equiv="expires" content="0" />
<link href="../css/Public.css" rel="stylesheet" type="text/css" /> 
<script type="text/javascript" src="../Include/Pub.js"></script>
<script type="text/javascript" src="../Include/TipBox.js"></script>
<script type="text/javascript">
function Copy(Obj){ 
	var clipBoardContent = $Get(Obj).value; 
	$Get(Obj).select();
	window.clipboardData.setData("Text",clipBoardContent); 
	//alert("复制成功!"); 
	new x.creat(1, 41, 5, 10, '复制成功!');
} 
function MakeCode(){
	//大类===========================================================
	var bClass = document.getElementsByName("BigClass");
	var bClassArr = bClass.length;
	var bClassStr = "";
	for (i = 0;i < bClassArr;i++){
		if(bClass[i].checked == true){
			bClassStr = bClassStr + bClass[i].value + ",";
		}
	}
	bClassStr = bClassStr.substring(0,bClassStr.length-1)
	if(bClassStr){
		var bClassSql = " Or ClassNum = '"+ bClassStr +"'";
	}else{
		var bClassSql = "";
	}
	
	//代码===============================================================
	var Code2 = "";
	Code2 = "&lt;table width=\"100%\" border=\"0\" cellspacing=\"0\" cellpadding=\"0\">\n"+
			"&lt;form method=\"post\" action=\"\">\n"+
			"&lt;input name=\"ClassNum\" type=\"hidden\" value=\""+ bClassStr +"\" />\n"+
			"&lt;%\n"+
			"	Set rs = Server.CreateObject(\"ADODB.Recordset\")\n"+
			"	rs.open \"Select * From p8_Field Where 1=1 "+ bClassSql +" Order By id Asc\",Conn,1,1\n"+
			"	\n"+
			"		FieldStr = \"\"\n"+
			"		FieldId  = \"\"\n"+
			"\n"+
			"		Do While Not rs.Eof \n"+
			"			\n"+
			"			FieldId = FieldId & \",\" & rs(\"Variable\")\n"+
			"			\n"+
			"			If rs(\"MaxLen\") &lt;> 0 Then \n"+
			"				MaxLen = \" maxlength=\"\"\"& rs(\"MaxLen\") &\"\"\"\"\n"+
			"			End If\n"+
			"			\n"+
			"			'单行文本框 =============================================\n"+
			"			If rs(\"FieldType\") = \"text\" Then\n"+
			"				FieldStr = FieldStr & \"&lt;tr>&lt;td width=\"\"100\"\" height=\"\"30\"\" align=\"\"right\"\">\"& rs(\"FieldName\") &\"：&nbsp;&lt;/td>&lt;td colspan=\"\"6\"\">&lt;input name=\"\"\"& rs(\"Variable\") &\"\"\" type=\"\"text\"\" class=\"\"ipt4\"\" value=\"\"\"& rs(\"Content\") &\"\"\" id=\"\"\"& rs(\"Variable\") &\"\"\" style=\"\"width:\"& rs(\"Width\") &\";\"\" \"& MaxLen &\">&lt;/td>&lt;/tr>\"\n"+
			"			End If\n"+
			"			\n"+
			"			'多行文本框 =============================================\n"+
			"			If rs(\"FieldType\") = \"textarea\" Then\n"+
			"				FieldStr = FieldStr & \"&lt;tr>&lt;td width=\"\"100\"\" height=\"\"30\"\" align=\"\"right\"\">\"& rs(\"FieldName\") &\"：&nbsp;&lt;/td>&lt;td colspan=\"\"6\"\" style=\"\"padding-top:5px;\"\">&lt;textarea name=\"\"\"& rs(\"Variable\") &\"\"\" id=\"\"\"& rs(\"Variable\") &\"\"\" class=\"\"ipt3\"\" style=\"\"width:\"& rs(\"Width\") &\"px; height:\"& rs(\"Height\") &\"px;\"\">\"& rs(\"Content\") &\"&lt;/textarea>&lt;/td>&lt;/tr>\"\n"+
			"			End If\n"+
			"			\n"+
			"			'单选框 =================================================\n"+
			"			If rs(\"FieldType\") = \"radio\" Then\n"+
			"				FieldStr = FieldStr & \"&lt;tr>&lt;td width=\"\"100\"\" height=\"\"30\"\" align=\"\"right\"\">\"& rs(\"FieldName\") &\"：&nbsp;&lt;/td>&lt;td colspan=\"\"6\"\">\"\n"+
			"				\n"+
			"				Options = Split(rs(\"Options\"),chr(13))\n"+
			"				\n"+
			"				For j = 0 To Ubound(Options)\n"+
			"					FieldStr = FieldStr & \"&lt;input type=\"\"radio\"\" name=\"\"\"& rs(\"Variable\") &\"\"\" value=\"\"|\"& Replace(Trim(Options(j)),chr(10),\"\") &\"|\"\" />\"& Trim(Options(j)) &\" \" \n"+
			"				Next\n"+
			"				\n"+
			"				FieldStr = FieldStr & \"&lt;/td>&lt;/tr>\"\n"+
			"			End If\n"+
			"			\n"+
			"			'复选框 =================================================\n"+
			"			If rs(\"FieldType\") = \"checkbox\" Then\n"+
			"				FieldStr = FieldStr & \"&lt;tr>&lt;td width=\"\"100\"\" height=\"\"30\"\" align=\"\"right\"\">\"& rs(\"FieldName\") &\"：&nbsp;&lt;/td>&lt;td colspan=\"\"6\"\">\"\n"+
			"				\n"+
			"				Options = Split(rs(\"Options\"),chr(13))\n"+
			"				\n"+
			"				For j = 0 To Ubound(Options)\n"+
			"					FieldStr = FieldStr & \"&lt;input type=\"\"checkbox\"\" name=\"\"\"& rs(\"Variable\") &\"\"\" value=\"\"|\"& Replace(Trim(Options(j)),chr(10),\"\") &\"|\"\" />\"& Trim(Options(j)) &\" \" \n"+
			"				Next\n"+
			"				\n"+
			"				FieldStr = FieldStr & \"&lt;/td>&lt;/tr>\"\n"+
			"			End If\n"+
			"			\n"+
			"			'下拉框 =================================================\n"+
			"			If rs(\"FieldType\") = \"select\" Then\n"+
			"				FieldStr = FieldStr & \"&lt;tr>&lt;td width=\"\"100\"\" height=\"\"30\"\" align=\"\"right\"\">\"& rs(\"FieldName\") &\"：&nbsp;&lt;/td>&lt;td colspan=\"\"6\"\">\"\n"+
			"				\n"+
			"				Options = Split(rs(\"Options\"),chr(13))\n"+
			"				\n"+
			"				FieldStr = FieldStr & \"&lt;select style=\"\"width:\"& rs(\"Width\") &\";\"\" name=\"\"\"& rs(\"Variable\") &\"\"\">&lt;option value=\"\"\"\">&lt;/option>\"\n"+
			"				\n"+
			"				For j = 0 To Ubound(Options)\n"+
			"					FieldStr = FieldStr & \"&lt;option value=\"\"|\"& Replace(Trim(Options(j)),chr(10),\"\") &\"|\"\">\"& Trim(Options(j)) &\"&lt;/option>\" \n"+
			"				Next\n"+
			"				\n"+
			"				FieldStr = FieldStr & \"&lt;/select>\"\n"+
			"				\n"+
			"				FieldStr = FieldStr & \"&lt;/td>&lt;/tr>\"\n"+
			"			End If\n"+
			"		\n"+
			"		rs.MoveNext         \n"+
			"	Loop\n"+
			"	rs.Close\n"+
			"	Set rs = Nothing\n"+
			"	\n"+
			"	If Left(FieldId,1) = \",\" Then FieldId = Right(FieldId,Len(FieldId)-1)\n"+
			"\n"+
			"	Response.Write FieldStr & \"&lt;input type=\"\"hidden\"\" name=\"\"FieldId\"\" value=\"\"\"& FieldId &\"\"\" />\"\n"+
			"%&gt;\n"+
			"&lt;tr>\n"+
			"  &lt;td width=\"100\" height=\"30\" align=\"right\">&nbsp;&lt;/td>\n"+
			"  &lt;td colspan=\"6\">&lt;input type=\"submit\" name=\"Submit\" value=\"提交\" />&lt;/td>\n"+
			"&lt;/tr>\n"+
			"&lt;/form>\n"+
			"&lt;/table>\n"+
			"&lt;%\n"+
			"	CloseConn\n"+
			"%&gt;"

	Code2 = Code2.replace(new RegExp("1=1  Or "),""); 
	Code2 = Code2.replace(new RegExp(" Where 1=1  Order")," Order");
	Code2 = Code2.replace(new RegExp('&lt;', 'g'),'<');
	Code2 = Code2.replace(new RegExp('&gt;', 'g'),'>');
	
	
	$Get("Code2").value = Code2;  //写入代码
}
</script>
</head>

<body>
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
    <tr>
      <td bgcolor="#F8FBFE"><table border="0" cellspacing="10" cellpadding="0">
        <tr>
          <td width="80" height="30" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Code_News_List.asp';">文章列表</td>
          <td width="80" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Code_News_View.asp';">文章显示</td>
          <td width="80" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Code_Pic_List.asp';">图片列表</td>
          <td width="80" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Code_Pic_View.asp';">图片显示</td>
          <td width="80" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Code_Page.asp';">单页</td>
          <td width="80" align="center" class="Tab1_over">表单</td>
          <td width="80" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Code_Service.asp';">在线客服</td>
          <td width="80" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Code_User.asp';">登录框</td>
        </tr>
      </table>	  </td>
    </tr>
    <tr>
      <td bgcolor="#F8FBFE" style="padding:10px;"><table width="100%" border="0" cellpadding="10" cellspacing="1" bgcolor="#E4EDF9">
          <tr>
            <td bgcolor="#FFFFFF" style="line-height:160%;"><strong>字段说明：<br>
              </strong><span class="cGray">内容 &lt;%=rs(&quot;Content&quot;)%&gt;&nbsp;自定义字段 &lt;%=NewsField(&quot;变量名&quot;,rs(&quot;id&quot;))%&gt; </span><br>              
              <strong>其他说明：</strong><br>
              <span class="cGray">放置代码前，请保证需要放置代码的文件扩展名为.asp，如asp文件中包含该代码“&lt;%@LANGUAGE=&quot;VBSCRIPT&quot; CODEPAGE=&quot;936&quot;%&gt;”，请将其删除。</span></td>
          </tr>
      </table></td>
    </tr>
  </table>

<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">

<tr>
  <td width="74" height="30" align="center" bgcolor="#F8FBFE">&nbsp;&nbsp;&nbsp;&nbsp;分类：</td>
  <td bgcolor="#F8FBFE">
  <%
	Set rs2 = Server.Createobject("Adodb.RecordSet")
	rs2.open "Select ClassName,Num From p8_Class Where ClassType='表单' And ClassLevel=1 Order By id Desc",Conn,1,1
	
		Do While Not rs2.Eof 
			
			Response.Write "<input type=""radio"" name=""BigClass"" value="""& rs2("Num") &""" /><strong>"& rs2("ClassName") &"</strong>&nbsp;"

		rs2.MoveNext         
		Loop 

	rs2.Close
	Set rs2 = Nothing
  %>  </td>
</tr>
<tr>
  <td height="30" align="center" bgcolor="#F8FBFE">&nbsp;</td>
  <td bgcolor="#F8FBFE"><span style="padding:20px 0;">
    <input name="Submit2" type="submit" class="btn2" value="生成代码" onClick="MakeCode()" >
  </span></td>
</tr>
<tr>
  <td height="30" align="center" bgcolor="#F8FBFE">&nbsp;</td>
  <td bgcolor="#F8FBFE">&nbsp;</td>
</tr>
<tr>
  <td height="30" align="center" bgcolor="#F8FBFE">数据处理：</td>
  <td bgcolor="#F8FBFE"><textarea id="Code1" class="ipt3" style="width:550px; height:180px;" readonly="readonly">&lt;!--#include file="Include/Class_Conn.asp"--&gt;
&lt;!--#include file="Include/Class_Main.asp"--&gt;
&lt;%
	If Request("Submit")<>"" Then
		ClassNum   = Trim(Request.Form("ClassNum"))
		FieldId    = Trim(Request.Form("FieldId"))
		sFieldId   = Split(FieldId,",")'告诉应该接收哪些字段名
		
		For i = 0 To Ubound(sFieldId)
			FieldContent = FieldContent & "{$"& sFieldId(i) &"$}"& Request.Form(sFieldId(i)) &"{$/"& sFieldId(i) &"$}"
		Next

		If Instr(Replace(FieldContent,"$}{$/",""),"{$/")<=0 Then
			Response.Write "<script>alert(""不能提交空信息！"");window.history.back();</script>"
			Response.End()
		End If
		
		Set rs = Server.Createobject("Adodb.RecordSet")
		rs.open "Select id From p8_Class Where Num='"& ClassNum &"'",Conn,1,1
		
			If Not rs.Eof Then
				ClassID = rs("id")
			Else
				Response.Write "<script>alert(""参数错误！"");window.history.back();</script>"
				Response.End()
			End If

		rs.Close
		Set rs = Nothing

		Set rs=Server.CreateObject("Adodb.Recordset")
		rs.Open "Select * From p8_Form",Conn,1,3
		rs.AddNew

		rs("ClassID")      = ClassID
		rs("FieldContent") = FieldContent

		rs.Update
		rs.Close
		Set rs=Nothing
		Response.Write "<script>alert(""提交成功！"");window.location.href=window.location.href;</script>"
	End If
%&gt;</textarea>
      <br>
      <input name="Submit" type="button" class="btn3" onClick="Copy('Code1')" value="复制以上代码">
    将以上代码放到网页最顶部</td>
</tr>
<tr>
  <td height="30" align="center" bgcolor="#F8FBFE">表单代码：</td>
  <td bgcolor="#F8FBFE"><textarea id="Code2" class="ipt3" style="width:550px; height:280px;" readonly="readonly">&lt;table width="100%" border="0" cellspacing="0" cellpadding="0">
&lt;form method="post" action="">
&lt;input name="ClassNum" type="hidden" value="" />
&lt;%
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.open "Select * From p8_Field Where ClassNum = '' Order By id Asc",Conn,1,1
	
		FieldStr = ""
		FieldId  = ""

		Do While Not rs.Eof 
			
			FieldId = FieldId & "," & rs("Variable")
			
			If rs("MaxLen") &lt;> 0 Then 
				MaxLen = " maxlength="""& rs("MaxLen") &""""
			End If
			
			'单行文本框 =============================================
			If rs("FieldType") = "text" Then
				FieldStr = FieldStr & "&lt;tr>&lt;td width=""100"" height=""30"" align=""right"">"& rs("FieldName") &"：&nbsp;&lt;/td>&lt;td colspan=""6"">&lt;input name="""& rs("Variable") &""" type=""text"" class=""ipt4"" value="""& rs("Content") &""" id="""& rs("Variable") &""" style=""width:"& rs("Width") &";"" "& MaxLen &">&lt;/td>&lt;/tr>"
			End If
			
			'多行文本框 =============================================
			If rs("FieldType") = "textarea" Then
				FieldStr = FieldStr & "&lt;tr>&lt;td width=""100"" height=""30"" align=""right"">"& rs("FieldName") &"：&nbsp;&lt;/td>&lt;td colspan=""6"" style=""padding-top:5px;"">&lt;textarea name="""& rs("Variable") &""" id="""& rs("Variable") &""" class=""ipt3"" style=""width:"& rs("Width") &"px; height:"& rs("Height") &"px;"">"& rs("Content") &"&lt;/textarea>&lt;/td>&lt;/tr>"
			End If
			
			'单选框 =================================================
			If rs("FieldType") = "radio" Then
				FieldStr = FieldStr & "&lt;tr>&lt;td width=""100"" height=""30"" align=""right"">"& rs("FieldName") &"：&nbsp;&lt;/td>&lt;td colspan=""6"">"
				
				Options = Split(rs("Options"),chr(13))
				
				For j = 0 To Ubound(Options)
					FieldStr = FieldStr & "&lt;input type=""radio"" name="""& rs("Variable") &""" value=""|"& Replace(Trim(Options(j)),chr(10),"") &"|"" />"& Trim(Options(j)) &" " 
				Next
				
				FieldStr = FieldStr & "&lt;/td>&lt;/tr>"
			End If
			
			'复选框 =================================================
			If rs("FieldType") = "checkbox" Then
				FieldStr = FieldStr & "&lt;tr>&lt;td width=""100"" height=""30"" align=""right"">"& rs("FieldName") &"：&nbsp;&lt;/td>&lt;td colspan=""6"">"
				
				Options = Split(rs("Options"),chr(13))
				
				For j = 0 To Ubound(Options)
					FieldStr = FieldStr & "&lt;input type=""checkbox"" name="""& rs("Variable") &""" value=""|"& Replace(Trim(Options(j)),chr(10),"") &"|"" />"& Trim(Options(j)) &" " 
				Next
				
				FieldStr = FieldStr & "&lt;/td>&lt;/tr>"
			End If
			
			'下拉框 =================================================
			If rs("FieldType") = "select" Then
				FieldStr = FieldStr & "&lt;tr>&lt;td width=""100"" height=""30"" align=""right"">"& rs("FieldName") &"：&nbsp;&lt;/td>&lt;td colspan=""6"">"
				
				Options = Split(rs("Options"),chr(13))
				
				FieldStr = FieldStr & "&lt;select style=""width:"& rs("Width") &";"" name="""& rs("Variable") &""">&lt;option value="""">&lt;/option>"
				
				For j = 0 To Ubound(Options)
					FieldStr = FieldStr & "&lt;option value=""|"& Replace(Trim(Options(j)),chr(10),"") &"|"">"& Trim(Options(j)) &"&lt;/option>" 
				Next
				
				FieldStr = FieldStr & "&lt;/select>"
				
				FieldStr = FieldStr & "&lt;/td>&lt;/tr>"
			End If
		
		rs.MoveNext         
	Loop
	rs.Close
	Set rs = Nothing
	
	If Left(FieldId,1) = "," Then FieldId = Right(FieldId,Len(FieldId)-1)

	Response.Write FieldStr & "&lt;input type=""hidden"" name=""FieldId"" value="""& FieldId &""" />"
%&gt;
&lt;tr>
  &lt;td width="100" height="30" align="right">&nbsp;&lt;/td>
  &lt;td colspan="6">&lt;input type="submit" name="Submit" value="提交" />&lt;/td>
&lt;/tr>
&lt;/form>
&lt;/table>
&lt;%
	CloseConn
%&gt;
未选择表单分类
</textarea>
      <br>
      <input name="Submit" type="button" class="btn3" onClick="Copy('Code2')" value="复制以上代码">
    将以上代码放到&lt;/html&gt;后面</td>
</tr>
<tr>
  <td height="30" align="center" bgcolor="#F8FBFE">&nbsp;</td>
  <td bgcolor="#F8FBFE">&nbsp;</td>
</tr>
</table>

</body>
</html>
<%
	CloseConn
%>