<!--#include file="../../Include/Class_Conn.asp"-->
<!--#include file="../../Include/Class_Main.asp"-->
<!--#include file="../p8_Check.asp"-->
<!--#include file="../Include/upload_5xsoft.inc"-->
<%
	On Error Resume Next
	
	Dim i,ClassID,BigClass,SmallClass,PicName,Pic,Content,sFieldId,FieldContent
	Set upload = new upload_5xSoft
	
	ClassID  = Trim(upload.Form("ClassID"))
	PicName  = Trim(upload.Form("PicName"))
	Pic      = Trim(upload.Form("Pic"))
	Content  = Trim(upload.Form("Content"))
	
	'�����Զ����ֶ� =================================================================================
	FieldId    = Trim(upload.Form("FieldId"& ClassID &""))
	sFieldId   = Split(FieldId,",")'����Ӧ�ý�����Щ�ֶ���
	
	For i = 0 To Ubound(sFieldId)
		TagName    = Replace(sFieldId(i),"Fz_FieldId"& ClassID & "_" ,"") '���ʱ�����ʶ��
		FieldContent = FieldContent & "{$"& TagName &"$}"& upload.Form(sFieldId(i)) &"{$/"& TagName &"$}"
	Next
	'/�����Զ����ֶ� =================================================================================

	If ClassID="" Then
		Response.Write "<script>alert(""��ѡ���࣡"");window.history.back();</script>"
		Response.End()
	End If

	If PicName="" Then
		Response.Write "<script>alert(""����д���ƣ�"");window.history.back();</script>"
		Response.End()
	End If

	'ȡ�ô����С��ID
	Set rs = Server.Createobject("Adodb.RecordSet")
	rs.open "Select ClassLevel,ParentID From p8_Class Where id="& ClassID &"",Conn,1,1
	
		If Not rs.Eof Then
			If rs("ClassLevel") = 1 Then '���ѡ��ķ�����һ������ֱ�����ô���
				BigClass = ClassID
			End If
			
			If rs("ClassLevel") = 2 Then '���ѡ��ķ����Ƕ����������һ������ID
				BigClass   = rs("ParentID")
				SmallClass = ClassID
			End If
		End If

	rs.Close
	Set rs = Nothing

	Set rs=Server.CreateObject("Adodb.Recordset")
	rs.Open "Select * From p8_Pic",Conn,1,3
	rs.AddNew
	
	rs("BigClass")     = BigClass
	rs("SmallClass")   = SmallClass
	rs("PicName")      = PicName
	rs("Pic")          = Pic
	rs("Content")      = Content
	rs("Admin")        = Request.Cookies("Admin")("s_Name")
	rs("FieldContent") = FieldContent
	
	'�ϴ�ͼƬ======================================================================================================
	Function MakedownName()
	Dim fname
	fname = Now()
	fname = Replace(fname,"-","")
	fname = replace(fname,"/","")
	fname = replace(fname,".","") 	
	fname = Replace(fname," ","") 
	fname = Replace(fname,":","")
	fname = Replace(fname,"PM","")
	fname = Replace(fname,"AM","")
	fname = Replace(fname,"����","")
	fname = Replace(fname,"����","")
	fname = Int(fname) + Int((10-1+1)*Rnd + 1)
	MakedownName = fname
	End Function 
    
	Set FSO = CreateObject(FsoName)
	
	
	SavePath  = SiteDir & "UpFile/" & Year(Now) & Right("0"&Month(Now),2) & "/"
	SavePath2 = Replace(SiteDir,"/","\") & "UpFile\"
	If Not (FSO.FolderExists(Server.MapPath(SavePath2 & Year(Now) & Right("0"&Month(Now),2)))) Then
		FSO.CreateFolder(Server.MapPath(SavePath2 & Year(Now) & Right("0"&Month(Now),2)))
	End If
	Set FSO=Nothing

	If Err Then
		err.Clear
		Set Conn = Nothing
		Response.Write "���顰��վ�������еġ���վ����Ŀ¼���Ƿ���д��ȷ��"
		Response.End
	End If

	Set file = upload.file("SmallPic")
	If file.FileSize>0 Then
		imgtype=Lcase(Mid(file.FileName,Instrrev(file.FileName,".")+1))
		If  imgtype="gif" or imgtype="jpg" or imgtype="bmp" or imgtype="png" Then
			newname = MakedownName()&"."&mid(file.FileName,InStrRev(file.FileName, ".")+1)
			file.SaveAs Server.MapPath(SavePath & newname)
			rs("SmallPic")= SavePath & newname
		End If
	End If
	Set file=Nothing
	
	'/�ϴ�ͼƬ=====================================================================================================

	rs.Update
	rs.Close
	Set rs=Nothing
	CloseConn
	Response.Redirect "Pic_Add.asp?ClassID="& ClassID &"&Tip=��ӳɹ���"
	Response.End()
%>
