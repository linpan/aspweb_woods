<!--#include file="../../Include/Class_Conn.asp"-->
<!--#include file="../../Include/Class_Main.asp"-->
<% Super=1 'ֻ����������Ա���ʸ�ҳ %>
<!--#include file="../p8_Check.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�����б� - ���ݵ���</title>
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
	
	//��̬����=======================================================
	var ClassVar = $Get("ClassVar");
	var ClassVarStr1 = "";
	var ClassVarStr2 = "";
	if(ClassVar.checked == true){
		ClassVarStr1 = "	ClassID = Replace_Text(Request.QueryString(\"ClassID\"))\n"+
						"\n"+
						"	If ClassID <> \"\" Then\n"+
						"		If Not isNumeric(ClassID) Then\n"+
						"			Response.Write \"��������\"\n"+
						"			Response.End()\n"+
						"		End If\n"+
						"		DtClassSql = \"And (BigClass = \"& ClassID &\" Or SmallClass = \"& ClassID &\")\"\n"+
						"	End If\n"+
						"\n"
		ClassVarStr2 = " \"& DtClassSql &\" "
		bClassSql = "";
		sClassSql = "";
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
		Code2 = "&lt;%\n"+ ClassVarStr1 +
				"	Set rs = Server.CreateObject(\"ADODB.Recordset\")\n"+
				"	rs.open \"Select * From p8_News Where 1=1 "+ ClassVarStr2 + bClassSql + sClassSql +" Order By "+ PaixuStr +" Desc\",Conn,1,1\n"+
				"	\n"+
				"	n = 1\n"+
				"	rs.PageSize = "+ PageCount +" 'ÿҳ��¼��\n"+
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
		Code2 = "&lt;%\n"+ ClassVarStr1 +
				"	Set rs = Server.CreateObject(\"ADODB.Recordset\")\n"+
				"	rs.open \"Select Top "+ RowCount +" * From p8_News Where 1=1 "+ ClassVarStr2 + bClassSql + sClassSql +" Order By "+ PaixuStr +" Desc\",Conn,1,1\n"+
				"	\n"+
				"	n = 1\n"+
				"	If rs.RecordCount>0 Then\n"+
				"		Do While Not rs.Eof And n <= "+ RowCount +"\n"+
				"%&gt;\n";
	}

	
	Code2 = Code2.replace(new RegExp("1=1  Or "),"");
	Code2 = Code2.replace(new RegExp("1=1  BigClass"),"BigClass");
	Code2 = Code2.replace(new RegExp(" Where 1=1  Order")," Order"); 
	Code2 = Code2.replace(new RegExp("&lt;"),"<"); 
	Code2 = Code2.replace(new RegExp("&gt;"),">"); 
	
	$Get("Code2").value = Code2;  //д�����
	
	
	//alert(Code2);
}
function DoClass(){
	if($Get("ClassVar").checked == true){
		$Get("ClassDiv").disabled = "disabled";
		$Get("ClassVarTr").style.display = "";
	}else{
		$Get("ClassDiv").disabled = "";
		$Get("ClassVarTr").style.display = "none";
	}
}
</script>
</head>

<body>
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
    <tr>
      <td bgcolor="#F8FBFE"><table border="0" cellspacing="10" cellpadding="0">
        <tr>
          <td width="80" height="30" align="center" class="Tab1_over">�����б�</td>
          <td width="80" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Code_News_View.asp';">������ʾ</td>
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
              <strong>ѭ�����ݷ�����</strong><br>
              <span class="cGray">&lt;a href=&quot;NewsView.asp?id=&lt;%=rs(&quot;id&quot;)%&gt;&quot; target=&quot;_blank&quot;&gt;&lt;font color=&quot;&lt;%=rs(&quot;TitleColor&quot;)%&gt;&quot;&gt;&lt;%rs(&quot;Title&quot;)%&gt;&lt;/font&gt;&lt;/a&gt;&lt;%=rs(&quot;AddDate&quot;)%&gt; (����NewsView.asp�ĳ������ʾҳ�ļ���) </span><br>              <strong>����˵����</strong><br>
              <span class="cGray">���ô���ǰ���뱣֤��Ҫ���ô�����ļ���չ��Ϊ.asp����asp�ļ��а����ô��롰&lt;%@LANGUAGE=&quot;VBSCRIPT&quot; CODEPAGE=&quot;936&quot;%&gt;�����뽫��ɾ����</span></td>
          </tr>
      </table></td>
    </tr>
  </table>

<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">

<tr>
  <td height="30" align="center" bgcolor="#F8FBFE">&nbsp;&nbsp;&nbsp;&nbsp;���ࣺ</td>
  <td bgcolor="#F8FBFE">
  <span id="ClassDiv">
  <%
	Set rs2 = Server.Createobject("Adodb.RecordSet")
	rs2.open "Select id,ClassName From p8_Class Where ClassType='����' And ClassLevel=1 Order By id Desc",Conn,1,1
	
		Do While Not rs2.Eof 
			
			Response.Write "<input type=""checkbox"" name=""BigClass"" value="""& rs2("id") &""" /><strong>"& rs2("ClassName") &"</strong>&nbsp;"
			
			Set rs3 = Server.CreateObject("ADODB.RecordSet")
			rs3.open "Select id,ClassName From p8_Class Where ClassType='����' And ClassLevel=2 And ParentID="& rs2("id") &" Order By id Desc",Conn,1,1
			
			Do While Not rs3.Eof 						
				Response.Write "<input type=""checkbox"" name=""SmallClass"" value="""& rs3("id") &""" />"& rs3("ClassName") &"&nbsp;"
				rs3.MoveNext         
			Loop 
			
			rs3.Close
			Set rs3 = Nothing

		rs2.MoveNext         
		Loop 

	rs2.Close
	Set rs2 = Nothing
  %></span>  <input type="checkbox" id="ClassVar" name="ClassVar" value="1" onClick="DoClass()"><span class="cGreen">��̬����</span><br>
<div id="ClassVarTr" style="display:none; margin-top:10px; border:1px solid #F9D751; line-height:180%; padding:10px; background-color:#FEF9E2; color:#A86500;">
 ��̬�����ǽ��������һ���ļ��У�ͨ����ַ��������ʾ��Ӧ�������ݡ���ʽ�磺NewsList.asp?ClassID=3�����С�NewsList.asp���ɸĳ�����ļ�������3���ĳɶ�Ӧ����ID�����ձ����£�������Ϊ��Ӧ����ID��<br>
<%
Set rs2 = Server.Createobject("Adodb.RecordSet")
rs2.open "Select id,ClassName From p8_Class Where ClassType='����' And ClassLevel=1 Order By id Desc",Conn,1,1

	Do While Not rs2.Eof 
		
		Response.Write "<strong>"& rs2("ClassName") &"</strong>("& rs2("id") &")&nbsp;"
		
		Set rs3 = Server.CreateObject("ADODB.RecordSet")
		rs3.open "Select id,ClassName From p8_Class Where ClassType='����' And ClassLevel=2 And ParentID="& rs2("id") &" Order By id Desc",Conn,1,1
		
		Do While Not rs3.Eof 						
			Response.Write ""& rs3("ClassName") &"("& rs3("id") &")&nbsp;"
			rs3.MoveNext         
		Loop 
		
		rs3.Close
		Set rs3 = Nothing

	rs2.MoveNext         
	Loop 

rs2.Close
Set rs2 = Nothing
%>
</div>
  </td>
</tr>
<tr>
  <td height="30" align="center" bgcolor="#F8FBFE">&nbsp;&nbsp;&nbsp;&nbsp;����</td>
  <td bgcolor="#F8FBFE"><input name="Paixu" type="radio" value="AddDate" checked>�������&nbsp;<input name="Paixu" type="radio" value="Hits">������</td>
</tr>
<tr>
  <td height="30" align="center" bgcolor="#F8FBFE">�Ƿ��ҳ��</td>
  <td bgcolor="#F8FBFE"><input name="isPage" type="radio" onClick="$Get('PageCountTr').style.display='';$Get('RowCountTr').style.display='none';$Get('PageCodeTr').style.display='';" value="1" checked>��&nbsp;<input name="isPage" type="radio" onClick="$Get('PageCountTr').style.display='none';$Get('RowCountTr').style.display='';$Get('PageCodeTr').style.display='none';" value="0">��</td>
</tr>
<tr id="PageCountTr">
  <td height="30" align="center" bgcolor="#F8FBFE">ÿҳ������</td>
  <td bgcolor="#F8FBFE"><input name="PageCount" type="text" class="ipt3" id="PageCount" value="10" size="10" maxlength="5"></td>
</tr>
<tr id="RowCountTr" style="display:none;">
  <td height="30" align="center" bgcolor="#F8FBFE">��ʾ������</td>
  <td bgcolor="#F8FBFE"><input name="RowCount" type="text" class="ipt3" id="RowCount" value="5" size="10" maxlength="5"></td>
</tr>
<tr>
  <td height="30" align="center" bgcolor="#F8FBFE">&nbsp;</td>
  <td bgcolor="#F8FBFE"><span style="padding:20px 0;">
    <input name="Submit2" type="submit" class="btn2" value="���ɴ���" onClick="MakeCode()" >
  </span></td>
</tr>
<tr>
  <td height="30" align="center" bgcolor="#F8FBFE">&nbsp;</td>
  <td bgcolor="#F8FBFE">&nbsp;</td>
</tr>
<tr>
  <td width="74" height="30" align="center" bgcolor="#F8FBFE">����������</td>
  <td bgcolor="#F8FBFE"><textarea id="Code1" class="ipt3" style="width:550px; height:40px;" readonly="readonly">&lt;!--#include file="Include/Class_Conn.asp"--&gt;
&lt;!--#include file="Include/Class_Main.asp"--&gt;</textarea>
    <br>
    <input name="Submit" type="button" class="btn3" onClick="Copy('Code1')" value="�������ϴ���">
    �����ϴ���ŵ���ҳ���(��ҳ����������ͬ���룬����Ҫ����)</td>
</tr>
<tr>
  <td height="30" align="center" bgcolor="#F8FBFE">��ʼ��ȡ��</td>
  <td bgcolor="#F8FBFE">
<textarea id="Code2" class="ipt3" style="width:550px; height:100px;" readonly="readonly">
&lt;%
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.open "Select * From p8_News Order By AddDate Desc",Conn,1,1

	n = 1
	rs.PageSize = 10 'ÿҳ��¼��
	tatalrecord = rs.RecordCount '�ܼ�¼��
	tatalpages  = rs.PageCount '��ҳ��
	page        = Request("page")
	
	If tatalpages = 0 Then tatalpages = 1
	If Not isNumeric(page) Or page="" Then page = 1
	If Cint(page) > Cint(tatalpages) Then page = tatalpages
	If tatalrecord <> 0 Then rs.AbsolutePage = page
	
	If tatalrecord>0 Then
		Do While Not rs.Eof And n <= rs.PageSize
%&gt;</textarea>
<br>
      <input name="Submit" type="button" class="btn3" onClick="Copy('Code2')" value="�������ϴ���">
      �����ϴ���ŵ�ѭ���б�ǰ</td>
</tr>
<tr>
  <td height="30" align="center" bgcolor="#F8FBFE">������ȡ��</td>
  <td bgcolor="#F8FBFE"><textarea id="Code3" class="ipt3" style="width:550px; height:100px;" readonly="readonly">
&lt;%
			n = n + 1
			rs.MoveNext
		Loop 
	End If
%&gt;</textarea>
      <br>
      <input name="Submit" type="button" class="btn3" onClick="Copy('Code3')" value="�������ϴ���">
    �����ϴ���ŵ�ѭ���б�ǰ</td>
</tr>
<tr id="PageCodeTr">
  <td height="30" align="center" bgcolor="#F8FBFE">��ҳ���룺</td>
  <td bgcolor="#F8FBFE"><textarea id="Code4" class="ipt3" style="width:550px; height:100px;" readonly="readonly">
&lt;%
If tatalpages>1 Then
	If page>1 Then
		PageHtml = "<a href=""?page="& page-1 &"&ClassID="& ClassID &""">��һҳ</a> "
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
	
	If Clng(page)<Clng(tatalpages) Then PageHtml = PageHtml & "<a href=""?page="& page+1 &"&ClassID="& ClassID &""">��һҳ</a> "
	
	PageHtml = PageHtml & "ת��<select onchange=""window.location.href= '?page='+ this.value + '&ClassID="& ClassID &"';"">"
	For p=1 To tatalpages
		If Clng(p) = Clng(page) Then
			PageHtml = PageHtml & "<option value="""& p &""" selected=""selected"">"& p &"</option>"
		Else
			PageHtml = PageHtml & "<option value="""& p &""">"& p &"</option>"
		End If
	Next
	PageHtml = PageHtml & "</select>ҳ"

	Response.Write PageHtml

End If
%&gt;
</textarea>
      <br>
      <input name="Submit" type="button" class="btn3" onClick="Copy('Code4')" value="�������ϴ���">
    ��ҳ���������ڽ�����ȡ���棬���ô����ɽ�������</td>
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