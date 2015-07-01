<!--#include file="../../Include/Class_Conn.asp"-->
<!--#include file="../../Include/Class_Main.asp"-->
<!--#include file="../p8_Check.asp"-->
<!--#include file="../Include/upload_5xsoft.inc"-->
<%
	Dim id,Page,s_ClassID,s_PicName,s_Content,s_Pic_px,i,ClassID,BigClass,SmallClass,PicName,Pic,Content,sFieldId,FieldContent
	Set upload = new upload_5xSoft
	
	id        = Trim(upload.Form("id"))
	Page      = Trim(upload.Form("Page"))
	s_ClassID = Trim(upload.Form("s_ClassID"))
	s_PicName = Trim(upload.Form("s_PicName"))
	s_Content = Trim(upload.Form("s_Content"))
	s_Pic_px  = Trim(upload.Form("s_Pic_px"))
	
	ClassID  = Trim(upload.Form("ClassID"))
	PicName  = Trim(upload.Form("PicName"))
	Pic      = Trim(upload.Form("Pic"))
	Content  = Trim(upload.Form("Content"))
	
	'接收自定义字段 =================================================================================
	FieldId    = Trim(upload.Form("FieldId"& ClassID &""))
	sFieldId   = Split(FieldId,",")'告诉应该接收哪些字段名

	For i = 0 To Ubound(sFieldId)
		TagName    = Replace(sFieldId(i),"Fz_FieldId"& ClassID & "_" ,"") '入库时清除标识符
		FieldContent = FieldContent & "{$"& TagName &"$}"& upload.Form(sFieldId(i)) &"{$/"& TagName &"$}"
	Next
	'/接收自定义字段 =================================================================================

	If ClassID="" Then
		Response.Write "<script>alert(""请选分类！"");window.history.back();</script>"
		Response.End()
	End If

	If PicName="" Then
		Response.Write "<script>alert(""请填写名称！"");window.history.back();</script>"
		Response.End()
	End If

	'取得大类和小类ID
	Set rs = Server.Createobject("Adodb.RecordSet")
	rs.open "Select ClassLevel,ParentID From p8_Class Where id="& ClassID &"",Conn,1,1
	
		If Not rs.Eof Then
			If rs("ClassLevel") = 1 Then '如果选择的分类是一级，则直接设置大类
				BigClass = ClassID
			End If
			
			If rs("ClassLevel") = 2 Then '如果选择的分类是二级，则查找一级分类ID
				BigClass   = rs("ParentID")
				SmallClass = ClassID
			End If
		End If

	rs.Close
	Set rs = Nothing

	Set rs=Server.CreateObject("Adodb.Recordset")
	rs.Open "Select * From p8_Pic Where id="&id,Conn,1,3
	
	rs("BigClass")     = BigClass
	rs("SmallClass")   = SmallClass
	rs("PicName")      = PicName
	rs("Pic")          = Pic
	rs("Content")      = Content
	rs("Admin")        = Request.Cookies("Admin")("s_Name")
	rs("FieldContent") = FieldContent
	
	'上传图片======================================================================================================
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
	fname = Replace(fname,"上午","")
	fname = Replace(fname,"下午","")
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

	Set file = upload.file("SmallPic")
	If file.FileSize>0 Then
		imgtype=Lcase(Mid(file.FileName,Instrrev(file.FileName,".")+1))
		If  imgtype="gif" or imgtype="jpg" or imgtype="bmp" or imgtype="png" Then
			newname = MakedownName()&"."&mid(file.FileName,InStrRev(file.FileName, ".")+1)
			
			Set fdel = CreateObject(FsoName)	'删除原图
			If Left(rs("SmallPic"),8)="/UpFile/"  Then   
				If (fdel.FileExists(Server.MapPath(rs("SmallPic")))) Then
					fdel.DeleteFile(Server.MapPath(rs("SmallPic")))
				End If 
			End If
			Set fdel=Nothing 
			
			file.SaveAs Server.MapPath(SavePath & newname)
			rs("SmallPic")= SavePath & newname
		End If
	End If
	Set file=Nothing
	'/上传图片=====================================================================================================

	rs.Update
	rs.Close
	Set rs=Nothing
	CloseConn
	Response.Redirect "Pic_List.asp?Tip=修改成功！&Page="& Page &"&ClassID="& s_ClassID &"&PicName="& s_PicName &"&Content="& s_Content &"&Pic_px="& s_Pic_px
	Response.End()
%>
