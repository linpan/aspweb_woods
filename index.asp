<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!--#include file="Include/Class_Conn.asp"-->
<!--#include file="Include/Class_Main.asp"-->
<%
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.open "Select Top 1 * From p8_Page Where ClassID in(41)",Conn,1,1
	
	If rs.Eof Then
		Response.Write "暂无内容！"
		Response.End()
	Else
		If rs("ClassID") <> "" Then
			Set rss = Server.CreateObject("ADODB.Recordset")
			rss.open "Select UserLimit From p8_Class Where id = " & rs("ClassID"),Conn,1,1	
			If Not rss.Eof Then
				UserLimit = rss("UserLimit")
			End If
			rss.Close
			Set rss = Nothing
		End If
	
		If UserLimit = "1" And Session("UserState") <> "1" Then
			Response.Write "<script>alert(""该内容只有正式会员才能查看\n\n请先登录或注册！"");window.location.href='"& SiteDir &"User/Login.asp';</script>"
			Response.End()
		End If
	End If
%>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>psd8模板_集团企业</title>
<link href="css/global.css" rel="stylesheet" type="text/css" />
<link href="css/index.css" rel="stylesheet" type="text/css" />
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
//-->
</script>
</head>

<body>
<!--#include file="top.asp"-->
<!--top-->
<div id="banner">
	<div id="con_tit"><img src="images/titnews.gif" width="164" height="14" class="left" /><a href="news.asp"><img src="images/more.gif" width="31" height="14" class="right" /></a></div>
</div>
<!--banner-->
<div id="con_bg">
	<div id="con">
    	<div id="con_l"><img src="images/about_tit.gif" width="239" height="30" /><p><em><img src="images/pic.gif" width="42" height="35" /></em><%=left(rs("Content"),50)%></p>
        	<ul>
            	<li><a href="qywh.asp">企业文化</a></li>
                <li><a href="qyry.asp">企业荣誉</a></li>
            </ul>
        </div>
        <div id="con_m">
    <%
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.open "Select Top 1 * From p8_News Where BigClass in(39) Order By AddDate Desc",Conn,1,1
	
	n = 1
	If rs.RecordCount>0 Then
		Do While Not rs.Eof And n <= 1
%>
        	<div id="news"><a href="news_view.asp?id=<%=rs("id")%>" target="_blank"><img src="<%=rs("SmallPic")%>" width="131" height="76" /></a><span class="left"><b><a href="news_view.asp?id=<%=rs("id")%>" target="_blank"><%=rs("Title")%></a></b><p>    <%=rs("Content")%></p></span></div>
            <%
			n = n + 1
			rs.MoveNext
		Loop 
	End If
%>
			<ul>
            <%
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.open "Select Top 3 * From p8_News Where BigClass in(37) Order By AddDate Desc",Conn,1,1
	
	n = 1
	If rs.RecordCount>0 Then
		Do While Not rs.Eof And n <= 3
%>
            	<li><a href="news_view.asp?id=<%=rs("id")%>" target="_blank" class="left"><%=rs("Title")%></a> <a href="news_view.asp?id=<%=rs("id")%>" target="_blank" class="right"><%=left(rs("AddDate"),10)%></a></li>
                <%
			n = n + 1
			rs.MoveNext
		Loop 
	End If
%>
            </ul>
   	  </div>
        <div id="con_r">
        	<div id="tit">
            	<ul>
                	<li id="one1" onmouseover="setTab('one',1,2)" class="hover">知名案例</li>
                    <li id="one2" onmouseover="setTab('one',2,2)">专题平速</li>
                </ul>
            </div>
            <div id="con_rm">
            <div  id="con_one_1" class="hover">
            	<ul>
                <%
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.open "Select Top 5 * From p8_News Where BigClass in(40) Order By AddDate Desc",Conn,1,1
	
	n = 1
	If rs.RecordCount>0 Then
		Do While Not rs.Eof And n <= 5
%>
                	<li><a href="news_view.asp?id=<%=rs("id")%>" target="_blank" class="left"><%=rs("Title")%></a></li>
                    <%
			n = n + 1
			rs.MoveNext
		Loop 
	End If
%>
                </ul>
            </div>
            <div  id="con_one_2" style="display:none">
            	<ul>
                	<%
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.open "Select Top 5 * From p8_News Where BigClass in(40) Order By AddDate Desc",Conn,1,1
	
	n = 1
	If rs.RecordCount>0 Then
		Do While Not rs.Eof And n <= 4
%>
                	<li><a href="news_view.asp?id=<%=rs("id")%>" target="_blank" class="left"><%=rs("Title")%></a></li>
                    <%
			n = n + 1
			rs.MoveNext
		Loop 
	End If
%>
                </ul>
            </div>
            </div>
        </div>
    </div>
</div>
<!--con-->
<div id="bottom_bg">
	<div id="bottom">
    	<span class="left"><a href="login.asp">会员登录</a> |  <a href="#">友情链接</a> |  <a href="contact.asp">联系方式</a></span>
        <span class="right">Copyright@2010 SHENGSHICHUANMEI GROUP.LTD,All Rights Reserved</span>
    </div>
</div>
</body>
</html>
<%
	CloseConn
%>