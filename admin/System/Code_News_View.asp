<!--#include file="../../Include/Class_Conn.asp"-->
<!--#include file="../../Include/Class_Main.asp"-->
<% Super=1 'ֻ����������Ա���ʸ�ҳ %>
<!--#include file="../p8_Check.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>������ʾ - ���ݵ���</title>
<script type="text/javascript">top.window.aTitle.innerText='���ݵ���'</script>
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
	//alert("���Ƴɹ�!"); 
	new x.creat(1, 41, 5, 10, '���Ƴɹ�!');
} 
function MakeCode(){
	//����===========================================================
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
	
	//С��===========================================================
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
	
	//�Ƿ��ҳ===========================================================
	var isPage = document.getElementsByName("isPage");
	var isPageStr = "";
	for(var i = 0;i < isPage.length;i++){
		if(isPage[i].checked == true){
			isPageStr = isPage[i].value;
			break;
		}
	}
	
	//����===========================================================
	var Paixu = document.getElementsByName("Paixu");
	var PaixuStr = "";
	for(var i = 0;i < Paixu.length;i++){
		if(Paixu[i].checked == true){
			PaixuStr = Paixu[i].value;
			break;
		}
	}
	
	//ÿҳ����===========================================================
	var PageCount = document.getElementById("PageCount").value;
	var re = new RegExp(/^(-|\+)?\d+$/);
	if(!re.test(PageCount) && $Get("PageCountTr").style.display!="none"){alert("ÿҳ��������Ϊ��ֵ");return false;}
	
	//��ʾ����===========================================================
	var RowCount = document.getElementById("RowCount").value;
	var re = new RegExp(/^(-|\+)?\d+$/);
	if(!re.test(RowCount) && $Get("RowCountTr").style.display!="none"){alert("��ʾ�кű���Ϊ��ֵ");return false;}
	
	//����===============================================================
	var Code2 = "";
	if(isPageStr == 1){ //���ѡ���˷�ҳ
		Code2 = "&lt;%\n"+
				"	Set rs = Server.CreateObject(\"ADODB.Recordset\")\n"+
				"	rs.open \"Select * From p8_News Where 1=1 "+ bClassSql + sClassSql +" Order By "+ PaixuStr +" Desc\",Conn,1,1\n"+
				"	\n"+
				"	n = 1\n"+
				"	rs.PageSize = 10 'ÿҳ��¼��\n"+
				"	tatalrecord = rs.RecordCount '�ܼ�¼��\n"+
				"	tatalpages  = rs.PageCount '��ҳ��\n"+
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
	
	$Get("Code2").value = Code2;  //д�����
	
	
	//alert(Code2);
}
</script>
</head>

<body>
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
    <tr>
      <td bgcolor="#F8FBFE"><table border="0" cellspacing="10" cellpadding="0">
        <tr>
          <td width="80" height="30" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Code_News_List.asp';">�����б�</td>
          <td width="80" align="center" class="Tab1_over">������ʾ</td>
          <td width="80" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Code_Pic_List.asp';">ͼƬ�б�</td>
          <td width="80" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Code_Pic_View.asp';">ͼƬ��ʾ</td>
          <td width="80" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Code_Page.asp';">��ҳ</td>
          <td width="80" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Code_Form.asp';">��</td>
          <td width="80" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Code_Service.asp';">���߿ͷ�</td>
          <td width="80" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Code_User.asp';">��¼��</td>
        </tr>
      </table>	  </td>
    </tr>
    <tr>
      <td bgcolor="#F8FBFE" style="padding:10px;"><table width="100%" border="0" cellpadding="10" cellspacing="1" bgcolor="#E4EDF9">
          <tr>
            <td bgcolor="#FFFFFF" style="line-height:160%;"><strong>�ֶ�˵����<br>
              </strong><span class="cGray">���� &lt;%=rs(&quot;Title&quot;)%&gt;&nbsp;&nbsp;���� &lt;%=rs(&quot;AddDate&quot;)%&gt;&nbsp;&nbsp;��Դ &lt;%=rs(&quot;Source&quot;)%&gt; &nbsp;&nbsp;������ɫ &lt;%=rs(&quot;TitleColor&quot;)%&gt; &nbsp;&nbsp;������ַ &lt;%=Url%&gt;&nbsp;&nbsp;�ؼ��� &lt;%=rs(&quot;KeyWord&quot;)%&gt;&nbsp;&nbsp;����ͼƬ &lt;%=rs(&quot;SmallPic&quot;)%&gt;&nbsp;&nbsp;�������� &lt;%=rs(&quot;Content&quot;)%&gt;&nbsp;&nbsp;������� &lt;%=rs(&quot;Hits&quot;)%&gt;&nbsp;&nbsp;�Զ����ֶ� &lt;%=NewsField(&quot;������&quot;,rs(&quot;id&quot;))%&gt; </span><br>              
              <strong>����˵����</strong><br>
              <span class="cGray">���ô���ǰ���뱣֤��Ҫ���ô�����ļ���չ��Ϊ.asp����asp�ļ��а����ô��롰&lt;%@LANGUAGE=&quot;VBSCRIPT&quot; CODEPAGE=&quot;936&quot;%&gt;�����뽫��ɾ����</span></td>
          </tr>
      </table></td>
    </tr>
  </table>

<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
<tr>
  <td width="74" height="30" align="center" bgcolor="#F8FBFE">����ͨѶ��</td>
  <td bgcolor="#F8FBFE"><textarea id="Code1" class="ipt3" style="width:550px; height:260px;" readonly="readonly">&lt;!--#include file="Include/Class_Conn.asp"--&gt;
&lt;!--#include file="Include/Class_Main.asp"--&gt;
&lt;%
	id = Replace_Text(Request.QueryString("id"))
	
	If id = "" Then
		Response.Write "��������"
		Response.End()
	End If
	
	If Not isNumeric(id) Then
		Response.Write "��������"
		Response.End()
	End If
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.open "Select * From p8_News Where id = " & id,Conn,1,1

	If rs.Eof Then
		Response.Write "����Ϣ�����ڣ�"
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
		Response.Write "<script>alert(""������ֻ����ʽ��Ա���ܲ鿴\n\n���ȵ�¼��ע�ᣡ"");window.location.href='"& SiteDir &"User/Login.asp';</script>"
		Response.End()
	End If
%&gt;</textarea>
    <br>
    <input name="Submit" type="button" class="btn3" onClick="Copy('Code1')" value="�������ϴ���">
    �����ϴ���ŵ���ҳ���</td>
</tr>
<tr>
  <td height="30" align="center" bgcolor="#F8FBFE">�ر����ӣ�</td>
  <td bgcolor="#F8FBFE">
<textarea id="Code2" class="ipt3" style="width:550px; height:70px;" readonly="readonly">
&lt;%
	Conn.Execute = "Update p8_News Set Hits = "& Hits + 1 &" Where id = " & id
	CloseConn
%&gt;</textarea>
<br>
      <input name="Submit" type="button" class="btn3" onClick="Copy('Code2')" value="�������ϴ���">
      �����ϴ���ŵ�&lt;/html&gt;����</td>
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