<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!--#include file="Include/Class_Conn.asp"-->
<!--#include file="Include/Class_Main.asp"-->
<%
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
%>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>psd8模板_集团企业</title>
<link href="css/global.css" rel="stylesheet" type="text/css" />
<link href="css/Inside.css" rel="stylesheet" type="text/css" />
<!--[if IE 6]>
	<script type="text/javascript" src="js/DD_belatedPNG.js"></script>
	
	<script type="text/javascript">
	 DD_belatedPNG.fix('img, #con_tit');
	</script>
<![endif]-->
<!--[if lte IE 6]>
<style type="text/css">
body { behavior:url("js/csshover.htc"); }
</style>
<![endif]-->
<script language=JavaScript type="text/javascript" src="js/tab.js"></script>
<script type="text/javascript">
<!--
function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>
</head>

<body onload="MM_preloadImages('images/menu_1.gif','images/menu1_1.gif','images/menu2_1.gif','images/menu3_1.gif','images/menu4_1.gif','images/menu5_1.gif','images/menu6_1.gif')">
<!--#include file="top.asp"-->
<!--top-->
<div id="banner2">
	<div id="con_tit"><img src="images/con_titbg1.png" width="999" height="27" /></div>
</div>
<!--banner-->
<div class="con_bg">
	<div class="con_bg1">
		<div class="con">
			<div class="con_l left">
            	<h2><img src="images/tit2.gif" width="146" height="15" /></h2>
                <ul>
                	<li><a href="contact.asp">联系方式</a></li>
                    <li><a href="messages.asp" class="hover">在线留言</a></li>
                </ul>
              <div class="contact"><img src="images/contacttit.gif" width="224" height="71" /><p>客服热线：0739-5230760 5332900 付先生：13886760438<br />
陈小姐：13632458845<br />传  真：020-87219960<br />E-mail：shijicheng@163.com <br />公司地址：湖南省大祥区云溪北路商务综合楼310室 </p></div>
                <a href="team.asp"><img src="images/team.gif" width="224" height="47" /></a>
            </div>
    		<div class="con_r right">
            	<div class="title">
                	<span class="left"><img src="images/title6.gif" width="138" height="14" /></span>
                  <p class="right">您现在的位置：<a href="index.asp">首页</a>--<a href="contact.asp">联系我们</a>--<a href="messages.asp">在线留言</a></p>
  	  	  	  	</div>
              	<div class="main8">
              		<p><b class="cred">留言请注意：</b>为了给您提供更好、更及时的答复，请认真填写以上内容带”*”选项为必填项。</p>
                    <div class="box1_m">
                    	<table width="100%" border="0" cellspacing="0" cellpadding="0">
<form method="post" action="">
<input name="ClassNum" type="hidden" value="1172094940PRH0X" />
<%
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.open "Select * From p8_Field Where ClassNum = '1172094940PRH0X' Order By id Asc",Conn,1,1
	
		FieldStr = ""
		FieldId  = ""

		Do While Not rs.Eof 
			
			FieldId = FieldId & "," & rs("Variable")
			
			If rs("MaxLen") <> 0 Then 
				MaxLen = " maxlength="""& rs("MaxLen") &""""
			End If
			
			'单行文本框 =============================================
			If rs("FieldType") = "text" Then
				FieldStr = FieldStr & "<tr><td width=""100"" height=""30"" align=""right"">"& rs("FieldName") &"：&nbsp;</td><td colspan=""6""><input name="""& rs("Variable") &""" type=""text"" class=""ipt4"" value="""& rs("Content") &""" id="""& rs("Variable") &""" style=""width:"& rs("Width") &";"" "& MaxLen &"></td></tr>"
			End If
			
			'多行文本框 =============================================
			If rs("FieldType") = "textarea" Then
				FieldStr = FieldStr & "<tr><td width=""100"" height=""30"" align=""right"">"& rs("FieldName") &"：&nbsp;</td><td colspan=""6"" style=""padding-top:5px;""><textarea name="""& rs("Variable") &""" id="""& rs("Variable") &""" class=""ipt3"" style=""width:"& rs("Width") &"px; height:"& rs("Height") &"px;"">"& rs("Content") &"</textarea></td></tr>"
			End If
			
			'单选框 =================================================
			If rs("FieldType") = "radio" Then
				FieldStr = FieldStr & "<tr><td width=""100"" height=""30"" align=""right"">"& rs("FieldName") &"：&nbsp;</td><td colspan=""6"">"
				
				Options = Split(rs("Options"),chr(13))
				
				For j = 0 To Ubound(Options)
					FieldStr = FieldStr & "<input type=""radio"" name="""& rs("Variable") &""" value=""|"& Replace(Trim(Options(j)),chr(10),"") &"|"" />"& Trim(Options(j)) &" " 
				Next
				
				FieldStr = FieldStr & "</td></tr>"
			End If
			
			'复选框 =================================================
			If rs("FieldType") = "checkbox" Then
				FieldStr = FieldStr & "<tr><td width=""100"" height=""30"" align=""right"">"& rs("FieldName") &"：&nbsp;</td><td colspan=""6"">"
				
				Options = Split(rs("Options"),chr(13))
				
				For j = 0 To Ubound(Options)
					FieldStr = FieldStr & "<input type=""checkbox"" name="""& rs("Variable") &""" value=""|"& Replace(Trim(Options(j)),chr(10),"") &"|"" />"& Trim(Options(j)) &" " 
				Next
				
				FieldStr = FieldStr & "</td></tr>"
			End If
			
			'下拉框 =================================================
			If rs("FieldType") = "select" Then
				FieldStr = FieldStr & "<tr><td width=""100"" height=""30"" align=""right"">"& rs("FieldName") &"：&nbsp;</td><td colspan=""6"">"
				
				Options = Split(rs("Options"),chr(13))
				
				FieldStr = FieldStr & "<select style=""width:"& rs("Width") &";"" name="""& rs("Variable") &"""><option value=""""></option>"
				
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

	Response.Write FieldStr & "<input type=""hidden"" name=""FieldId"" value="""& FieldId &""" />"
%>
<tr>
  <td width="200" height="30" align="right">&nbsp;</td>
  <td colspan="6" class="annu"><input type="submit" name="Submit" value="确认留言" />
                    	<input type="reset" name="button2" id="button2" value="重新填写" /></td>
</tr>
</form>
</table>
<%
	CloseConn
%>
                    </div>
   	  		  	</div>
        	</div>
            <div class="clear"></div>
		</div>
    </div>
</div>
<!--con-->
<div class="bottom_jb"></div>
<div id="bottom_bg">
	<div id="bottom">
    	<span class="left"><a href="login.asp">会员登录</a> |  <a href="#">友情链接</a> |  <a href="contact.asp">联系方式</a></span>
        <span class="right">Copyright@2010 SHENGSHICHUANMEI GROUP.LTD,All Rights Reserved</span>
    </div>
</div>
</body>
</html>
