<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!--#include file="Include/Class_Conn.asp"-->
<!--#include file="Include/Class_Main.asp"-->
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
<div id="banner4">
	<div id="con_tit"><img src="images/con_titbg1.png" width="999" height="27" /></div>
</div>
<!--banner-->
<div class="con_bg">
	<div class="con_bg1">
		<div class="con">
			<div class="con_l left">
            	<h2><img src="images/tit4.gif" width="110" height="15" /></h2>
                <ul>
                	<li><a href="news.asp?ClassID=5" class="hover">企业新闻</a></li>
                    <li><a href="news.asp?ClassID=37">行业新闻</a></li>
                </ul>
              <div class="contact"><img src="images/contacttit.gif" width="224" height="71" /><p>客服热线：0739-5230760 5332900 付先生：13886760438<br />
陈小姐：13632458845<br />传  真：020-87219960<br />E-mail：shijicheng@163.com <br />公司地址：湖南省大祥区云溪北路商务综合楼310室 </p></div>
                <a href="team.html"><img src="images/team.gif" width="224" height="47" /></a>
            </div>
    		<div class="con_r right">
            	<div class="title">
                	<span class="left"><img src="images/title4.gif" width="194" height="14" /></span>
                    <p class="right">您现在的位置：<a href="index.asp">首页</a>--<a href="news.asp">新闻中心</a>--<a href="news.asp">企业新闻</a></p>
  	  	  	  </div>
                <div class="main5">
                <%
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.open "Select Top 1 * From p8_News Where BigClass in(45) Order By AddDate Desc",Conn,1,1
	
	n = 1
	If rs.RecordCount>0 Then
		Do While Not rs.Eof And n <= 1
%>
                	<div id="news"><a href="news_view.asp?id=<%=rs("id")%>" target="_blank"><img src="<%=rs("SmallPic")%>" width="131" height="76" /></a><span><b><a href="news_view.asp?id=<%=rs("id")%>" target="_blank"><%=rs("Title")%></a> </b>
           	      <p><%=rs("Content")%><a href="news_view.asp?id=<%=rs("id")%>" target="_blank" class="cred">>>查看详细</a></p>
   	        </span></div>
            <%
			n = n + 1
			rs.MoveNext
		Loop 
	End If
%><div class="clear"></div>
            		<ul>
                    <%
	ClassID = Replace_Text(Request.QueryString("ClassID"))

	If ClassID <> "" Then
		If Not isNumeric(ClassID) Then
			Response.Write "参数错误！"
			Response.End()
		End If
		DtClassSql = "And (BigClass = "& ClassID &" Or SmallClass = "& ClassID &")"
	End If

	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.open "Select * From p8_News Where 1=1  "& DtClassSql &"  Order By AddDate Desc",Conn,1,1
	
	n = 1
	rs.PageSize = 12 '每页记录数
	tatalrecord = rs.RecordCount '总记录数
	tatalpages  = rs.PageCount '总页数
	page        = Request("page")
	
	If tatalpages = 0 Then tatalpages = 1
	If Not isNumeric(page) Or page="" Then page = 1
	If Cint(page) > Cint(tatalpages) Then page = tatalpages
	If tatalrecord <> 0 Then rs.AbsolutePage = page
	
	If tatalrecord>0 Then
		Do While Not rs.Eof And n <= rs.PageSize
%>
                    	<li><a href="news_view.asp?id=<%=rs("id")%>" target="_blank" class="left"><%=rs("Title")%></a> <a href="news_view.asp?id=<%=rs("id")%>" target="_blank" class="right"><%=left(rs("AddDate"),10)%></a></li>
                        <%
			n = n + 1
			rs.MoveNext
		Loop 
	End If
%>
                    </ul>
                    <div class="order" style="margin-top:20px;"><%
If tatalpages>1 Then
	If page>1 Then
		PageHtml = "<a href=""?page="& page-1 &"&ClassID="& ClassID &""">上一页</a> "
	End If
	
	For k=page-4 To page+4
		If k>0 Then
			If Clng(k)<>Clng(page) Then
				PageHtml = PageHtml & "<a href=""?page="& k &"&ClassID="& ClassID &""">"& k &"</a> "
			Else
				PageHtml = PageHtml & "<span>"& k &"</span> "
			End If
		End If
		If k=tatalpages Then Exit For
	Next
	
	If tatalpages - page > 3 Then PageHtml = PageHtml & "... "
	
	If Clng(page)<Clng(tatalpages) Then PageHtml = PageHtml & "<a href=""?page="& page+1 &"&ClassID="& ClassID &""">下一页</a> "
	
	PageHtml = PageHtml & "转到<select onchange=""window.location.href= '?page='+ this.value + '&ClassID="& ClassID &"';"">"
	For p=1 To tatalpages
		If Clng(p) = Clng(page) Then
			PageHtml = PageHtml & "<option value="""& p &""" selected=""selected"">"& p &"</option>"
		Else
			PageHtml = PageHtml & "<option value="""& p &""">"& p &"</option>"
		End If
	Next
	PageHtml = PageHtml & "</select>页"

	Response.Write PageHtml

End If
%>
</div><div class="clear"></div>
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
