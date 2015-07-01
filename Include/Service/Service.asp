<!--#include file="../Class_Conn.asp"-->
<!--#include file="../Class_Main.asp"-->
<%
	Dim ServiceSwitch,ServiceID,ServiceCode,i,Im
	Set rs=Server.CreateObject("Adodb.Recordset")
	rs.Open "Select Top 1 ServiceSwitch,ServiceID,ServiceCode From p8_Config Order By id Desc",Conn,1,3
	
	If Not rs.Eof Then
		ServiceSwitch = rs("ServiceSwitch")
		ServiceID     = rs("ServiceID")
		ServiceCode   = rs("ServiceCode")
	End If

	rs.Close
	Set rs=Nothing
	
	If Cint(ServiceSwitch) = 1 Then
%>
document.writeln("<style type=\'text/css\'>");
document.writeln("	.ser_head {background:url(<%=SiteDir%>Include/Service/images/img3-5_2.gif) no-repeat;}");
document.writeln("	.info {padding-bottom: 10px; padding-left: 0px; padding-right: 0px; background: url(<%=SiteDir%>Include/Service/images/img3-5_3.gif) repeat-y; padding-top: 5px}");
document.writeln("	.down_kefu {width: 157px; background: url(<%=SiteDir%>Include/Service/images/img3-5_4.gif) no-repeat; height: 8px}");
document.writeln("	.Obtn {margin-top: 45px; width: 32px; background: url(<%=SiteDir%>Include/Service/images/img3-5_1.gif) no-repeat; float: left; height: 139px; margin-left: -5px}");
document.writeln("	.qqtable span {padding-bottom: 5px; line-height: 20px; padding-left: 0px; width: 100px; padding-right: 0px; color: #ff6600; font-size: 13px; font-weight: bold; padding-top: 5px}");
document.writeln("	.qqtable a,.qqtable a:hover {text-decoration: none}");
document.writeln("</style>");
document.writeln("<div id=\'flashzoo\' onmouseover=\'toBig()\' onmouseout=\'toSmall()\'>");
document.writeln("<table style=\'float:left\' border=\'0\' cellspacing=\'0\' cellpadding=\'0\' width=\'157\'>");
document.writeln("	<tr><td class=\'ser_head\' height=\'39\' valign=\'top\'>&nbsp;</td></tr>");
document.writeln("	<tr>");
document.writeln("	<td class=\'info\' valign=\'top\' height=\'140\'>");
document.writeln("		<table class=\'qqtable\' border=\'0\' cellspacing=\'0\' cellpadding=\'0\' width=\'120\' align=\'center\'>");
document.writeln("		<tr><td height=\'5\'></td></tr>");
<%
	ServiceID = Split(ServiceID,chr(13))
	
	For i = 0 To Ubound(ServiceID)
		
		If Instr(ServiceID(i),"=") Then
			Im = Split(Replace(ServiceID(i),chr(10),""),"=")

			If Lcase(Trim(Im(0))) = "qq" And Trim(Im(1))<>"" Then
				Response.Write "document.writeln(""<tr><td height=\'30\' align=\'middle\'><a target=\'_blank\' href=\'http://wpa.qq.com/msgrd?v=3&uin="& Trim(Im(1)) &"&site=qq&menu=yes\'><img border=\'0\' src=\'http://wpa.qq.com/pa?p=2:"& Trim(Im(1)) &":41\' alt=\'点击这里给我发消息\' title=\'点击这里给我发消息\'></a></td></tr>"");"
			End If
			
			If Lcase(Trim(Im(0))) = "msn" And Trim(Im(1))<>"" Then
				Response.Write "document.writeln(""<tr><td height=\'30\' align=\'middle\'><a href=\'msnim:chat?contact="& Trim(Im(1)) &"\' target=\'_blank\'><img border=\'0\' src=\'"& SiteDir &"include/service/images/msn.png\' /></a></td></tr>"");"
			End If
			
			If Lcase(Trim(Im(0))) = "旺旺" And Trim(Im(1))<>"" Then
				Response.Write "document.writeln(""<tr><td height=\'30\' align=\'middle\'><a target=\'_blank\' href=\'http://amos1.taobao.com/msg.ww?v=2&uid="& Trim(Im(1)) &"&s=1\' ><img border=\'0\' src=\'http://amos1.taobao.com/online.ww?v=2&uid="& Trim(Im(1)) &"&s=1\' alt=\'点击这里给我发消息\' title=\'点击这里给我发消息\' /></a></td></tr>"");"
			End If
			
			If Lcase(Trim(Im(0))) = "skype" And Trim(Im(1))<>"" Then
				Response.Write "document.writeln(""<tr><td height=\'30\' align=\'middle\'><a href=\'skype:"& Trim(Im(1)) &"?call\' on-click=\'return skypeCheck();\'><img src=http://mystatus.skype.com/smallclassic/"& Trim(Im(1)) &" style=\'border: none;\' alt=\'点击这里给我发消息\' title=\'点击这里给我发消息\' /></a></td></tr>"");"
			End If
			
			If Lcase(Trim(Im(0))) = "百度hi" And Trim(Im(1))<>"" Then
				Response.Write "document.writeln(""<tr><td height=\'30\' align=\'middle\'><a href=\'baidu://message/?id="& Trim(Im(1)) &"\'><img border=\'0\' src=\'http://tieba.baidu.com/tb/img/hi/hiOnline.gif\' alt=\'点击这里给我发消息\' title=\'点击这里给我发消息\'></a></td></tr>"");"
			End If
			
		End If
	Next
%>

document.writeln("		<tr><td align=\'middle\'>&nbsp;</td></tr>");
document.writeln("		</table>");
document.writeln("	</td></tr>");
document.writeln("	<tr><td class=\'down_kefu\' valign=\'top\'></td></tr>");
document.writeln("</table>");
document.writeln("<div class=\'Obtn\'></div>");
document.writeln("</div>");

<%
	ServiceCode = Replace(ServiceCode,chr(13),"")
	ServiceCode = Replace(ServiceCode,chr(10),"")
	ServiceCode = Replace(ServiceCode,"""","\""")
	ServiceCode = Replace(ServiceCode,"'","\'")
	Response.Write "document.writeln("""& ServiceCode &""");"
%>


Code8_Service = function (id,_top,_left){
	var me=id.charAt?document.getElementById(id):id, d1=document.body, d2=document.documentElement;
	d1.style.height=d2.style.height='100%';me.style.top=_top?_top+'px':0;me.style.left=_left+"px";
	me.style.position='absolute';
	setInterval(function (){me.style.top=parseInt(me.style.top)+(Math.max(d1.scrollTop,d2.scrollTop)+_top-parseInt(me.style.top))*0.1+'px';},10+parseInt(Math.random()*20));
	return arguments.callee;
};
window.onload=function (){
	Code8_Service
	('flashzoo',100,-152)
}

lastScrollY=0; 

var InterTime = 1;
var maxWidth=-1;
var minWidth=-152;
var numInter = 8;

var BigInter ;
var SmallInter ;

var o =  document.getElementById("flashzoo");
var i = parseInt(o.style.left);
function Big()
{
	if(parseInt(o.style.left)<maxWidth)
	{
		i = parseInt(o.style.left);
		i += numInter;	
		o.style.left=i+"px";	
		if(i==maxWidth)
			clearInterval(BigInter);
	}
}
function toBig()
{
	clearInterval(SmallInter);
	clearInterval(BigInter);
		BigInter = setInterval(Big,InterTime);
}
function Small()
{
	if(parseInt(o.style.left)>minWidth)
	{
		i = parseInt(o.style.left);
		i -= numInter;
		o.style.left=i+"px";
		
		if(i==minWidth)
			clearInterval(SmallInter);
	}
}
function toSmall()
{
	clearInterval(SmallInter);
	clearInterval(BigInter);
	SmallInter = setInterval(Small,InterTime);
	
}
<%End If%>