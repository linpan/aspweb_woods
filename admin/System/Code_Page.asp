<!--#include file="../../Include/Class_Conn.asp"-->
<!--#include file="../../Include/Class_Main.asp"-->
<% Super=1 'ֻ����������Ա���ʸ�ҳ %>
<!--#include file="../p8_Check.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��ҳ - ���ݵ���</title>
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
		var bClassSql = " Or ClassID in("+ bClassStr +")";
	}else{
		var bClassSql = "";
	}
	
	//����===============================================================
	var Code2 = "";
	Code2 = ""+
			"&lt;!--#include file=\"Include/Class_Conn.asp\"--&gt;\n"+
			"&lt;!--#include file=\"Include/Class_Main.asp\"--&gt;\n"+
			"&lt;%\n"+
			"	Set rs = Server.CreateObject(\"ADODB.Recordset\")\n"+
			"	rs.open \"Select Top 1 * From p8_Page Where 1=1 "+ bClassSql +"\",Conn,1,1\n"+
			"	\n"+
			"	If rs.Eof Then\n"+
			"		Response.Write \"�������ݣ�\"\n"+
			"		Response.End()\n"+
			"	Else\n"+		
			"		If rs(\"ClassID\") <> \"\" Then\n"+
			"			Set rss = Server.CreateObject(\"ADODB.Recordset\")\n"+
			"			rss.open \"Select UserLimit From p8_Class Where id = \" & rs(\"ClassID\"),Conn,1,1	\n"+
			"			If Not rss.Eof Then\n"+
			"				UserLimit = rss(\"UserLimit\")\n"+
			"			End If\n"+
			"			rss.Close\n"+
			"			Set rss = Nothing\n"+
			"		End If\n"+
			"	\n"+
			"		If UserLimit = \"1\" And Session(\"UserState\") <> \"1\" Then\n"+
			"			Response.Write \"&lt;script>alert(\"\"������ֻ����ʽ��Ա���ܲ鿴\\n\\n���ȵ�¼��ע�ᣡ\"\");window.location.href='\"& SiteDir &\"User/Login.asp';&lt;/script>\"\n"+
			"			Response.End()\n"+
			"		End If\n"+
			"	End If\n"+
			"%&gt;\n";

	Code2 = Code2.replace(new RegExp("1=1  Or "),""); 
	Code2 = Code2.replace(new RegExp(" Where 1=1  Order")," Order"); 
	Code2 = Code2.replace(new RegExp('&lt;', 'g'),'<');
	Code2 = Code2.replace(new RegExp('&gt;', 'g'),'>');
	
	$Get("Code1").value = Code2;  //д�����
}
</script>
</head>

<body>
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
    <tr>
      <td bgcolor="#F8FBFE"><table border="0" cellspacing="10" cellpadding="0">
        <tr>
          <td width="80" height="30" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Code_News_List.asp';">�����б�</td>
          <td width="80" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Code_News_View.asp';">������ʾ</td>
          <td width="80" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Code_Pic_List.asp';">ͼƬ�б�</td>
          <td width="80" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Code_Pic_View.asp';">ͼƬ��ʾ</td>
          <td width="80" align="center" class="Tab1_over">��ҳ</td>
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
              </strong><span class="cGray">���� &lt;%=rs(&quot;Content&quot;)%&gt;&nbsp;�Զ����ֶ� &lt;%=NewsField(&quot;������&quot;,rs(&quot;id&quot;))%&gt; </span><br>              
              <strong>����˵����</strong><br>
              <span class="cGray">���ô���ǰ���뱣֤��Ҫ���ô�����ļ���չ��Ϊ.asp����asp�ļ��а����ô��롰&lt;%@LANGUAGE=&quot;VBSCRIPT&quot; CODEPAGE=&quot;936&quot;%&gt;�����뽫��ɾ����</span></td>
          </tr>
      </table></td>
    </tr>
  </table>

<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">

<tr>
  <td width="74" height="30" align="center" bgcolor="#F8FBFE">&nbsp;&nbsp;&nbsp;&nbsp;���ࣺ</td>
  <td bgcolor="#F8FBFE">
  <%
	Set rs2 = Server.Createobject("Adodb.RecordSet")
	rs2.open "Select id,ClassName From p8_Class Where ClassType='��ҳ' And ClassLevel=1 Order By id Desc",Conn,1,1
	
		Do While Not rs2.Eof 
			
			Response.Write "<input type=""radio"" name=""BigClass"" value="""& rs2("id") &""" /><strong>"& rs2("ClassName") &"</strong>&nbsp;"

		rs2.MoveNext         
		Loop 

	rs2.Close
	Set rs2 = Nothing
  %>  </td>
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
  <td height="30" align="center" bgcolor="#F8FBFE">����ͨѶ��</td>
  <td bgcolor="#F8FBFE"><textarea id="Code1" class="ipt3" style="width:550px; height:180px;" readonly="readonly">&lt;!--#include file="Include/Class_Conn.asp"--&gt;
&lt;!--#include file="Include/Class_Main.asp"--&gt;
&lt;%
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.open "Select Top 1 * From p8_Page",Conn,1,1

	If rs.Eof Then
		Response.Write "�������ݣ�"
		Response.End()
	End If
%&gt;</textarea>
      <br>
      <input name="Submit" type="button" class="btn3" onClick="Copy('Code1')" value="�������ϴ���">
    �����ϴ���ŵ���ҳ���</td>
</tr>
<tr>
  <td height="30" align="center" bgcolor="#F8FBFE">�ر����ӣ�</td>
  <td bgcolor="#F8FBFE"><textarea id="Code2" class="ipt3" style="width:550px; height:70px;" readonly="readonly">
&lt;%
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