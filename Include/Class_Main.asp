<%
Dim cfg_rs,SiteDir,FsoName,UserCheck,UserLgErr,UserLgLock,SysLgErr,SysLgLock,SmtpEmail,SmtpUser,SmtpPass,SmtpServer

Set cfg_rs = Server.CreateObject("Adodb.Recordset")
cfg_rs.Open "Select Top 1 SiteDir,FsoName,UserCheck,UserLgErr,UserLgLock,SysLgErr,SysLgLock,SmtpEmail,SmtpUser,SmtpPass,SmtpServer From p8_Config Order By id Asc",Conn,1,1
If Not cfg_rs.Eof Then
	SiteDir    = cfg_rs("SiteDir")
	FsoName    = cfg_rs("FsoName")
	UserCheck  = cfg_rs("UserCheck")
	UserLgErr  = cfg_rs("UserLgErr")
	UserLgLock = cfg_rs("UserLgLock")
	SysLgErr   = cfg_rs("SysLgErr")
	SysLgLock  = cfg_rs("SysLgLock")
	SmtpEmail  = cfg_rs("SmtpEmail")
	SmtpUser   = cfg_rs("SmtpUser")
	SmtpPass   = cfg_rs("SmtpPass")
	SmtpServer = cfg_rs("SmtpServer")
Else
	Response.Write "网站参数表无法读取！"
	Response.End()
End If
cfg_rs.Close
Set cfg_rs=Nothing

'//发送邮件函数
Sub SendEmail(MailBox,MailBody,MailSubject)
	Cls_Mailname = SmtpUser
	Cls_Mailpass = SmtpPass
	Cls_Mailform = SmtpEmail
	Cls_Mailsmtp = SmtpServer
	
	Set Jmail=server.createobject("Jmail.Message")
	Jmail.Charset = "gb2312"
	JMail.ContentType = "text/html"
	Jmail.Silent = true
	Jmail.Priority = 3
	Jmail.MailServerUserName = Cls_Mailname  '有效电子邮件帐号
	Jmail.MailServerPassword = Cls_Mailpass  '有效电子邮件密码
	Jmail.From = Cls_Mailform    
	Jmail.FromName = Cls_WebName             '来自
	Jmail.Subject = MailSubject              '邮件标题
	
	Jmail.AddRecipient  ""&MailBox&""        '收件人的邮件地址  
	
	Jmail.Body = MailBody     '邮件内容       
	Jmail.Send(Cls_Mailsmtp)  'smtp服务器地址
	Set Jmail=nothing
End Sub

' ============================================   
' 格式化时间(显示)   
' 参数：n_Flag   
' 1:"yyyy-mm-dd hh:mm:ss"   
' 2:"yyyy-mm-dd"   
' 3:"hh:mm:ss"   
' 4:"yyyy年mm月dd日"   
' 5:"yyyymmdd"   
' 6:"yyyymmddhhmmss"    
' ============================================   
Function cTime(Str,i)   
 	If isDate(Str) = False Then Exit Function 
	Dim Str_Y,Str_M,Str_D,Str_H,Str_Me,Str_S
	Str_Y  = Year(Str)
	Str_M  = Month(Str)
	Str_D  = Day(Str)
	Str_H  = Hour(Str)
	Str_Me = Minute(Str)
	Str_S  = Second(Str)
	If Len(Str_M)  = 1 Then Str_M  = "0" & Str_M
	If Len(Str_D)  = 1 Then Str_D  = "0" & Str_D
	If Len(Str_H)  = 1 Then Str_H  = "0" & Str_H
	If Len(Str_Me) = 1 Then Str_Me = "0" & Str_Me
	If Len(Str_S)  = 1 Then Str_S  = "0" & Str_S
	
	Select Case i
		Case 1 : cTime = Str_Y & "年" & Str_M & "月" & Str_D & "日&nbsp;" & Str_H &"时" & Str_Me & "分" & Str_S & "秒"
		Case 2 : cTime = Str_Y & "-" & Str_M & "-" & Str_D & "&nbsp;" & Str_H &":" & Str_Me & ":" & Str_S
		Case 3 : cTime = Str_M & "." & Str_D
		Case 4 : cTime = Str_M & "月" & Str_D & "日"
		Case 5 : cTime = Str_Y & "." & Str_M & "." & Str_D
	End Select
End Function

Function Html2Txt(Str)
	Dim re
	Str=Lcase(Str)
	
	Set re = new RegExp
	re.IgnoreCase=True
	re.Global=True
	re.Pattern="(\<.[^\<]*\>)"
	Str=re.Replace(Str,"")
	re.Pattern="(\<\/[^\<]*\>)"
	Str=re.Replace(Str,"")
	
	Html2Txt = Str
	Set re = Nothing
	Set Str = Nothing
End Function

'记录历史操作
Sub History(His_Class,His_Name,His_ID)
	Dim His_User
	His_User = Request.Cookies("Admin")("s_User")
	
	If His_Name<>"" Then
		Set His_rs = Server.CreateObject("Adodb.Recordset")
		His_rs.Open "Select His_Class,His_Name,His_ID,His_User,His_Hit From p8_History Where His_Class = '"& His_Class &"' And His_Name = '"& His_Name &"' And His_ID = '"& His_ID &"' And His_User = '"& His_User &"'",Conn,1,3
		
		If His_rs.Eof Then
			His_rs.AddNew
			His_rs("His_Class") = His_Class
			His_rs("His_Name")  = His_Name
			His_rs("His_ID")    = His_ID
			His_rs("His_User")  = His_User
		Else
			His_rs("His_Hit") = Clng(His_rs("His_Hit"))+1
		End If
		
		His_rs.Update
		His_rs.Close
		Set His_rs=Nothing
	End If
End Sub

'生成标识码
Function MakeNum()
	Dim fname
	fname = Now()
	fname = Replace(fname,"-","")
	fname = Replace(fname," ","") 
	fname = Replace(fname,":","")
	fname = Replace(fname,"PM","")
	fname = Replace(fname,"AM","")
	fname = Replace(fname,"上午","")
	fname = Replace(fname,"下午","")
	fname = fname & Genkey(5)
	fname = Right(fname,Len(fname)-2)
	MakeNum = fname
End Function

Function Genkey(digits)
	Dim char_array(36)
	char_array(0) = "0"
	char_array(1) = "1"
	char_array(2) = "2"
	char_array(3) = "3"
	char_array(4) = "4"
	char_array(5) = "5"
	char_array(6) = "6"
	char_array(7) = "7"
	char_array(8) = "8"
	char_array(9) = "9"
	char_array(10) = "A"
	char_array(11) = "B"
	char_array(12) = "C"
	char_array(13) = "D"
	char_array(14) = "E"
	char_array(15) = "F"
	char_array(16) = "G"
	char_array(17) = "H"
	char_array(18) = "I"
	char_array(19) = "J"
	char_array(20) = "K"
	char_array(21) = "L"
	char_array(22) = "M"
	char_array(23) = "N"
	char_array(24) = "O"
	char_array(25) = "P"
	char_array(26) = "Q"
	char_array(27) = "R"
	char_array(28) = "S"
	char_array(29) = "T"
	char_array(30) = "U"
	char_array(31) = "V"
	char_array(32) = "W"
	char_array(33) = "X"
	char_array(34) = "Y"
	char_array(35) = "Z"
	
	Randomize
	
	Do While len(output) < digits
		num = char_array(int((36 - 0 + 1) * rnd + 0))
		output = output & num
	Loop
	
	Genkey = output
End Function

'根据字段变量名取字段内容 - 文章
Function GetField(FieldVariable,NewsID)
	If NewsID Then
		Dim FieldStr,GF_rs
		Set GF_rs = Server.CreateObject("Adodb.Recordset")
		GF_rs.Open "Select FieldContent From p8_News Where id = "& Clng(NewsID),Conn,1,1
		
		If Not GF_rs.Eof Then
			FieldStr = GF_rs("FieldContent")
			GetField = GetContent(FieldStr,"{$"& FieldVariable &"$}","{$/"& FieldVariable &"$}")
		End If
		
		GF_rs.Close
		Set GF_rs=Nothing
	End If
End Function

'根据字段变量名取字段内容 - 图片
Function GetPicField(FieldVariable,PicID)
	If PicID Then
		Dim FieldStr,GF_rs
		Set GF_rs = Server.CreateObject("Adodb.Recordset")
		GF_rs.Open "Select FieldContent From p8_Pic Where id = "& Clng(PicID),Conn,1,1
		
		If Not GF_rs.Eof Then
			FieldStr = GF_rs("FieldContent")
			GetPicField = GetContent(FieldStr,"{$"& FieldVariable &"$}","{$/"& FieldVariable &"$}")
		End If
		
		GF_rs.Close
		Set GF_rs=Nothing
	End If
End Function

'根据字段变量名取字段内容 - 表单
Function GetFormField(FieldVariable,PicID)
	If PicID Then
		Dim FieldStr,GF_rs
		Set GF_rs = Server.CreateObject("Adodb.Recordset")
		GF_rs.Open "Select FieldContent From p8_Form Where id = "& Clng(PicID),Conn,1,1
		
		If Not GF_rs.Eof Then
			FieldStr = GF_rs("FieldContent")
			FieldStr = Replace(FieldStr,"}|","}")
			FieldStr = Replace(FieldStr,"|{","{")
			FieldStr = GetContent(FieldStr,"{$"& FieldVariable &"$}","{$/"& FieldVariable &"$}")
			GetFormField = FieldStr
		End If
		
		GF_rs.Close
		Set GF_rs=Nothing
	End If
End Function

 '根据首尾截取内容
Function GetContent(HTML,starcode,endcode)
	If Instr(HTML,starcode)>0 And Instr(HTML,endcode)>0 Then
		Dim StartPos,EndPos,Length
		StartPos=Instr(1,HTML,starcode)
		EndPos=Instr(StartPos,HTML,endcode)
		Length=EndPos-StartPos
		GetContent=Replace(Mid(HTML,StartPos,Length),starcode,"")
	Else
		GetContent=""
	End If
End Function

'读取文章自定义字段内容
Function NewsField(FieldName,id)
	If FieldName<>"" And id<>"" And isNumeric(id) Then
		Dim FieldStr,F_rs
		Set F_rs = Server.CreateObject("Adodb.Recordset")
		F_rs.Open "Select FieldContent From p8_News Where id = "& Clng(id),Conn,1,1
		
		If Not F_rs.Eof Then
			FieldStr  = F_rs("FieldContent")
			If FieldStr<>"" Then
				FieldStr  = Replace(FieldStr,"|,","")
				FieldStr  = Replace(FieldStr,"|","")
			End If
			NewsField = GetContent(FieldStr,"{$"& FieldName &"$}","{$/"& FieldName &"$}")
		End If
		
		F_rs.Close
		Set F_rs=Nothing
	End If
End Function

'读取图片自定义字段内容
Function PicField(FieldName,id)
	If FieldName<>"" And id<>"" And isNumeric(id) Then
		Dim FieldStr,F_rs
		Set F_rs = Server.CreateObject("Adodb.Recordset")
		F_rs.Open "Select FieldContent From p8_Pic Where id = "& Clng(id),Conn,1,1
		
		If Not F_rs.Eof Then
			FieldStr  = F_rs("FieldContent")
			If FieldStr<>"" Then
				FieldStr  = Replace(FieldStr,"|,","")
				FieldStr  = Replace(FieldStr,"|","")
			End If
			PicField = GetContent(FieldStr,"{$"& FieldName &"$}","{$/"& FieldName &"$}")
		End If
		
		F_rs.Close
		Set F_rs=Nothing
	End If
End Function
%>