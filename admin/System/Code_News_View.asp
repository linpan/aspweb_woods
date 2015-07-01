<!--#include file="../../Include/Class_Conn.asp"-->
<!--#include file="../../Include/Class_Main.asp"-->
<% Super=1 '只允许超级管理员访问该页 %>
<!--#include file="../p8_Check.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>文章显示 - 数据调用</title>
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
		var bClassSql = " Or BigClass in("+ bClassStr +")";
	}else{
		var bClassSql = "";
	}
	
	//小类===========================================================
	var sClass = document.getElementsByName("SmallClass");
	var sClassArr = sClass.length;
	var sClassStr = "";
	for (i = 0;i < sClassArr;i++){
		if(sClass[i].checked == true){
			sClassStr = sClassStr + sClass[i].value + ",";
		}
	}
	sClassStr = sClassStr.substring(0,sClassStr.length-1)
	if(sClassStr){
		var sClassSql = " Or SmallClass in("+ sClassStr +")";
	}else{
		var sClassSql = "";
	}
	
	//是否分页===========================================================
	var isPage = document.getElementsByName("isPage");
	var isPageStr = "";
	for(var i = 0;i < isPage.length;i++){
		if(isPage[i].checked == true){
			isPageStr = isPage[i].value;
			break;
		}
	}
	
	//排序===========================================================
	var Paixu = document.getElementsByName("Paixu");
	var PaixuStr = "";
	for(var i = 0;i < Paixu.length;i++){
		if(Paixu[i].checked == true){
			PaixuStr = Paixu[i].value;
			break;
		}
	}
	
	//每页数量===========================================================
	var PageCount = document.getElementById("PageCount").value;
	var re = new RegExp(/^(-|\+)?\d+$/);
	if(!re.test(PageCount) && $Get("PageCountTr").style.display!="none"){alert("每页数量必须为数值");return false;}
	
	//显示行数===========================================================
	var RowCount = document.getElementById("RowCount").value;
	var re = new RegExp(/^(-|\+)?\d+$/);
	if(!re.test(RowCount) && $Get("RowCountTr").style.display!="none"){alert("显示行号必须为数值");return false;}
	
	//代码===============================================================
	var Code2 = "";
	if(isPageStr == 1){ //如果选择了分页
		Code2 = "&lt;%\n"+
				"	Set rs = Server.CreateObject(\"ADODB.Recordset\")\n"+
				"	rs.open \"Select * From p8_News Where 1=1 "+ bClassSql + sClassSql +" Order By "+ PaixuStr +" Desc\",Conn,1,1\n"+
				"	\n"+
				"	n = 1\n"+
				"	rs.PageSize = 10 '每页记录数\n"+
				"	tatalrecord = rs.RecordCount '总记录数\n"+
				"	tatalpages  = rs.PageCount '总页数\n"+
				"	page        = Request(\"page\")\n"+
				"	\n"+
				"	If tatalpages = 0 Then tatalpages = 1\n"+
				"	If Not isNumeric(page) Or page=\"\" Then page = 1\n"+
				"	If Cint(page) > Cint(tatalpages) Then page = tatalpages\n"+
				"	If tatalrecord <> 0 Then rs.AbsolutePage = page\n"+
				"	\n"+
				"	If tatalrecord>0 Then\n"+
				"		Do While Not rs.Eof And n <= rs.PageSize\n"+
				"%&gt;\n";
	}else{
		Code2 = "&lt;%\n"+
				"	Set rs = Server.CreateObject(\"ADODB.Recordset\")\n"+
				"	rs.open \"Select Top "+ RowCount +" * From p8_News Where 1=1 "+ bClassSql + sClassSql +" Order By "+ PaixuStr +" Desc\",Conn,1,1\n"+
				"	\n"+
				"	n = 1\n"+
				"	If rs.RecordCount>0 Then\n"+
				"		Do While Not rs.Eof And n <= "+ RowCount +"\n"+
				"%&gt;\n";
	}

	
	Code2 = Code2.replace(new RegExp("1=1  Or "),""); 
	Code2 = Code2.replace(new RegExp(" Where 1=1  Order")," Order"); 
	Code2 = Code2.replace(new RegExp("&lt;"),"<"); 
	Code2 = Code2.replace(new RegExp("&gt;"),">"); 
	
	$Get("Code2").value = Code2;  //写入代码
	
	
	//alert(Code2);
}
</script>
</head>

<body>
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
    <tr>
      <td bgcolor="#F8FBFE"><table border="0" cellspacing="10" cellpadding="0">
        <tr>
          <td width="80" height="30" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Code_News_List.asp';">文章列表</td>
          <td width="80" align="center" class="Tab1_over">文章显示</td>
          <td width="80" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Code_Pic_List.asp';">图片列表</td>
          <td width="80" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Code_Pic_View.asp';">图片显示</td>
          <td width="80" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Code_Page.asp';">单页</td>
          <td width="80" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Code_Form.asp';">表单</td>
          <td width="80" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Code_Service.asp';">在线客服</td>
          <td width="80" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Code_User.asp';">登录框</td>
        </tr>
      </table>	  </td>
    </tr>
    <tr>
      <td bgcolor="#F8FBFE" style="padding:10px;"><table width="100%" border="0" cellpadding="10" cellspacing="1" bgcolor="#E4EDF9">
          <tr>
            <td bgcolor="#FFFFFF" style="line-height:160%;"><strong>字段说明：<br>
              </strong><span class="cGray">标题 &lt;%=rs(&quot;Title&quot;)%&gt;&nbsp;&nbsp;日期 &lt;%=rs(&quot;AddDate&quot;)%&gt;&nbsp;&nbsp;来源 &lt;%=rs(&quot;Source&quot;)%&gt; &nbsp;&nbsp;标题颜色 &lt;%=rs(&quot;TitleColor&quot;)%&gt; &nbsp;&nbsp;外链地址 &lt;%=Url%&gt;&nbsp;&nbsp;关键字 &lt;%=rs(&quot;KeyWord&quot;)%&gt;&nbsp;&nbsp;标题图片 &lt;%=rs(&quot;SmallPic&quot;)%&gt;&nbsp;&nbsp;文章内容 &lt;%=rs(&quot;Content&quot;)%&gt;&nbsp;&nbsp;浏览次数 &lt;%=rs(&quot;Hits&quot;)%&gt;&nbsp;&nbsp;自定义字段 &lt;%=NewsField(&quot;变量名&quot;,rs(&quot;id&quot;))%&gt; </span><br>              
              <strong>其他说明：</strong><br>
              <span class="cGray">放置代码前，请保证需要放置代码的文件扩展名为.asp，如asp文件中包含该代码“&lt;%@LANGUAGE=&quot;VBSCRIPT&quot; CODEPAGE=&quot;936&quot;%&gt;”，请将其删除。</span></td>
          </tr>
      </table></td>
    </tr>
  </table>

<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
<tr>
  <td width="74" height="30" align="center" bgcolor="#F8FBFE">数据通讯：</td>
  <td bgcolor="#F8FBFE"><textarea id="Code1" class="ipt3" style="width:550px; height:260px;" readonly="readonly">&lt;!--#include file="Include/Class_Conn.asp"--&gt;
&lt;!--#include file="Include/Class_Main.asp"--&gt;
&lt;%
	id = Replace_Text(Request.QueryString("id"))
	
	If id = "" Then
		Response.Write "参数错误！"
		Response.End()
	End If
	
	If Not isNumeric(id) Then
		Response.Write "参数错误！"
		Response.End()
	End If
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.open "Select * From p8_News Where id = " & id,Conn,1,1

	If rs.Eof Then
		Response.Write "该信息不存在！"
		Response.End()
	End If
	
	If rs("SmallClass") <> "" Then
		Set rss = Server.CreateObject("ADODB.Recordset")
		rss.open "Select UserLimit From p8_Class Where id = " & rs("SmallClass"),Conn,1,1	
		If Not rss.Eof Then
			UserLimit = rss("UserLimit")
		End If
		rss.Close
		Set rss = Nothing
	Else
		Set rsb = Server.CreateObject("ADODB.Recordset")
		rsb.open "Select UserLimit From p8_Class Where id = " & rs("BigClass"),Conn,1,1	
		If Not rsb.Eof Then
			UserLimit = rsb("UserLimit")
		End If
	End If

	If UserLimit = "1" And Session("UserState") <> "1" Then
		Response.Write "<script>alert(""该内容只有正式会员才能查看\n\n请先登录或注册！"");window.location.href='"& SiteDir &"User/Login.asp';</script>"
		Response.End()
	End If
%&gt;</textarea>
    <br>
    <input name="Submit" type="button" class="btn3" onClick="Copy('Code1')" value="复制以上代码">
    将以上代码放到网页最顶部</td>
</tr>
<tr>
  <td height="30" align="center" bgcolor="#F8FBFE">关闭连接：</td>
  <td bgcolor="#F8FBFE">
<textarea id="Code2" class="ipt3" style="width:550px; height:70px;" readonly="readonly">
&lt;%
	Conn.Execute = "Update p8_News Set Hits = "& Hits + 1 &" Where id = " & id
	CloseConn
%&gt;</textarea>
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