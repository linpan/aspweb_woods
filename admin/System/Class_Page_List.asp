<!--#include file="../../Include/Class_Conn.asp"-->
<!--#include file="../../Include/Class_Main.asp"-->
<% Super=1 'ֻ����������Ա���ʸ�ҳ %>
<!--#include file="../p8_Check.asp"-->
<%
'ɾ��============================================================================================
If Request.QueryString("DelID")<>"" Then
	id = Request("DelID")
	
	Set rs = Server.Createobject("Adodb.RecordSet")
	rs.open "Select id,Num From p8_Class Where id= " & id ,Conn,1,3
		
		conn.Execute "Delete From p8_Field Where ClassNum = '"& rs("Num") &"'" 'ɾ���Զ����ֶ�
		
		Set rs2 = Server.Createobject("Adodb.RecordSet") 'ɾ����ҳ����
		rs2.open "Select Pic From p8_Page Where ClassID= " & id ,Conn,1,3
		
		If Not rs2.Eof Then
			Pic = rs2("Pic")
			Set FSO = CreateObject(FsoName)
			If Instr(Pic,"|") Then
				Pic = Split(Pic,"|")
				For i=0 To Ubound(Pic)
					If Pic(i)<>"" Then
						If  FSO.FileExists(Server.MapPath(Pic(i))) Then
							FSO.Deletefile(server.MapPath(Pic(i)))
						End IF
					End If
				Next
			ElseIf Pic<>"" Then
				If FSO.FileExists(Server.MapPath(Pic)) Then
					FSO.Deletefile(server.MapPath(Pic))
				End IF
			End If
			
		rs2.Delete		
		End If
		
		rs2.Close
		Set rs2 = Nothing
		
	
	rs.Delete
	rs.Close
	Set rs = Nothing
	Response.Redirect "Class_Page_List.asp?Tip=ɾ���ɹ���"
	Response.End()
End If
'/ɾ��============================================================================================
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��Ŀ����</title>
<script type="text/javascript">top.window.aTitle.innerText='��Ŀ����'</script>
<script type="text/javascript" src="../Include/TipBox.js"></script>
<meta http-equiv="pragma" content="no-cache" />
<meta http-equiv="Cache-Control" content="no-cache, must-revalidate" />
<meta http-equiv="expires" content="Wed, 26 Feb 1997 08:21:57 GMT" />
<meta http-equiv="expires" content="0" />
<link href="../css/Public.css" rel="stylesheet" type="text/css" /> 
</head>
<body>
<%
	Dim Tip
	Tip = Request.QueryString("Tip")
	If Tip <> "" Then
		Response.Write "<script type=""text/javascript"">window.onload=function(){new x.creat(1, 41, 5, 10, '"& Tip &"');}</script>"
	End If
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td bgcolor="#F8FBFE"><table border="0" cellspacing="10" cellpadding="0">
      <tr>
        <td width="80" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Class_News_List.asp';">���·���</td>
        <td width="80" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Class_Pic_List.asp';">ͼƬ����</td>
        <td width="80" height="30" align="center" class="Tab1_over">��ҳ����</td>
        <td width="80" align="center" class="Tab1" onMouseOver="this.className='Tab1_over2'" onMouseOut="this.className='Tab1'" onClick="window.location.href='Class_Form_List.asp';">������</td>
      </tr>
    </table></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="1" cellpadding="5">
  <tr bgcolor="#F8FBFE">
    <td height="25" bgcolor="#F8FBFE">&nbsp;<span class="f14 cBlack">��ҳ</span>&nbsp;&nbsp;&nbsp;<a href="Class_Page_Add.asp">���ӵ�ҳ</a></td>
  </tr>
  
<%
	Dim rs,cot,n,ico
	
	Set rs = Server.CreateObject("ADODB.RecordSet")
	rs.open "Select id,ClassName,ParentID,ClassLevel From p8_Class Where ClassType='��ҳ' And ClassLevel=1 Order By id Desc",Conn,1,1
	
	cot = rs.RecordCount
	n   = 1
	

	Do While Not rs.Eof
		If cot = n Then
			ico = "background:url(../images/icon.gif) #F8FBFE 20px -44px no-repeat;"
		Else
			ico = "background:url(../images/icon.gif) #F8FBFE 20px -18px no-repeat;"
		End If
%>
		<tr style="<%=ico%>" onMouseOver="this.style.backgroundColor='#e7ffa6'" onMouseOut="this.style.backgroundColor='#F8FBFE'">
		  <td height="25" style="padding-left:60px;"><table width="100%" border="0" cellspacing="0" cellpadding="0">
			<tr>
			  <td width="200" class="cBlack"><a href="Class_Page_Edit.asp?id=<%=rs("id")%>" class="cBlack"><%=rs("ClassName")%></a></td>
			  <td><a href="Class_Page_Edit.asp?id=<%=rs("id")%>">�޸�</a>&nbsp;|&nbsp;<a href="javascript:if(confirm('ɾ���󲻿ɻָ����Ƿ������'))window.location.href='?DelID=<%=rs("id")%>';">ɾ��</a></td>
			</tr>
		  </table></td>
		</tr>
<%
		rs.MoveNext
		n = n + 1
	Loop 

%>
</table>
</body>
</html>
<%
	CloseConn
%>