<% 
'//-----------------------------------------地址栏参数合法性检测-----------------------------------------
Dim Psd8_SqlInjdata,Psd8_SQL_inj,Psd8_strtemp
Psd8_SqlInjdata = "exec | insert |select |delete |update |count |*|chr|mid|master|truncate|char|declare"
Psd8_SQL_inj = split(Psd8_SqlInjdata,"|")
If Request.QueryString<>"" Then
For Each SQL_Get In Request.QueryString
  For SQL_Data=0 To Ubound(Psd8_SQL_inj)
    if instr(Request.QueryString(SQL_Get),Psd8_SQL_inj(Sql_DATA))>0 Then
		Response.Write "非法参数！"
		Response.End()
    end if
  next
Next
End If

Psd8_strtemp=request.servervariables("server_name")&request.servervariables("url")&"?"&request.QueryString
Psd8_strtemp=lcase(Psd8_strtemp)
if instr(Psd8_strtemp,"select%20") or instr(Psd8_strtemp,"insert%20") or instr(Psd8_strtemp,"delete%20from") or instr(Psd8_strtemp,"count(") or instr(Psd8_strtemp,"drop%20table") or instr(Psd8_strtemp,"update%20") or instr(Psd8_strtemp,"truncate%20") or instr(Psd8_strtemp,"asc(") or instr(Psd8_strtemp,"mid(") or instr(Psd8_strtemp,"char(") or instr(Psd8_strtemp,"xp_cmdshell") or instr(Psd8_strtemp,"exec%20master") or instr(Psd8_strtemp,"net%20user") or instr(Psd8_strtemp,"'") or instr(Psd8_strtemp,"""") or instr(Psd8_strtemp,"“") or instr(Psd8_strtemp,"”") or instr(Psd8_strtemp,":") or instr(Psd8_strtemp,": ") or instr(Psd8_strtemp,";") or instr(Psd8_strtemp,"; ") or instr(Psd8_strtemp,"%27")  then
	Response.Write "非法参数！"
	Response.End()
end if

'//-----------------------------------------表单提交的数据合法性检测-----------------------------------------
'//替换非法数据的函数
Function Replace_Text(fString)
	If fString<>"" Then
		if isnull(fString) then
			Replace_Text=""
			exit function
		else
			fString=trim(fString)
			fString=replace(fString,"'","''")
			fString=replace(fString,";","；")
			fString=replace(fString,"--","—")
			fString=replace(fString,"=","")
			'fString=replace(fString,"and","")
			'fString=replace(fString,"or","")
			fString=replace(fString,"select","")
			fString=replace(fString,"insert","")
			fString=replace(fString,"exec","")
			fString=replace(fString,"delete","")
			fString=replace(fString,"update","")
			fString=replace(fString,"count","")
			fString=replace(fString,"mid","")
			fString=replace(fString,"truncate","")
			fString=replace(fString,"%","")
			fString=replace(fString,"chr","")
			fString=replace(fString,"master","")
			fString=replace(fString,"char","")
			fString=replace(fString,"declare","")
			fString=replace(fString,"*","")
			fString=replace(fString,"from","")
			fString=server.htmlencode(fString)
			Replace_Text=fString
		end if	
	End If
End function

'//检测非法数据的函数
Function SafeRequest(ParaName) 
	Dim ParaValue 
	ParaValue=Request(ParaName)
	if IsNumeric(ParaValue)  then
	SafeRequest=ParaValue
	exit Function
	
	else
	ParaValuetemp=lcase(ParaValue)
	tempvalue="select |insert |delete from|'|count(|drop table|update |truncate  |asc(|mid(|char(|xp_cmdshell|exec master|net localgroup administrators|net user| and|%20from|exec|select|delete|count|*|%|chr|mid|master|truncate|char|declare"
	temps=split(tempvalue,"|")
	for mycount=0 to ubound(temps)
	if  Instr(ParaValuetemp,trim(temps(mycount))) > 0 then
		Response.Write "非法参数！"
		Response.End()
	end if
	next
	SafeRequest=ParaValue
	end if
	End function
	
	'//检测非法数据的函数
	Function SafeRequestform(ParaName) 
	Dim ParaValue 
	ParaValue=request.form(ParaName)
	if IsNumeric(ParaValue)  then
	SafeRequestform=ParaValue
	exit Function
	else
	ParaValuetemp=lcase(ParaValue)
	tempvalue="select |insert |delete from|'|count(|drop table|update |truncate  |asc(|mid(|char(|xp_cmdshell|exec master|net localgroup administrators|net user|  and|%20from|exec|select|delete|count|*|%|chr|mid|master|truncate|char|declare"
	temps=split(tempvalue,"|")
	for mycount=0 to ubound(temps)
	if  Instr(ParaValuetemp,trim(temps(mycount))) > 0 then
		Response.Write "非法参数！"
		Response.End()
	end if
	next
	SafeRequestform=ParaValue
	end if
End function

'//-----------------------------------------是否是本站提交的数据检测-----------------------------------------
Sub Check_url()
	If Instr(Lcase(request.serverVariables("HTTP_REFERER")),Lcase(request.ServerVariables("SERVER_NAME")))=0 then
		Response.Write "非法来路！"
		Response.End()
	End if 
End Sub

'//-----------------------------------------字符转换-----------------------------------------
Function HTMLEncode(fString)
	If not isnull(fString) then
		fString = replace(fString, ">", "&gt;")
		fString = replace(fString, "<", "&lt;")
	
		fString = Replace(fString, CHR(32), "&nbsp;")
		fString = Replace(fString, CHR(9), "&nbsp;")
		fString = Replace(fString, CHR(34), "&quot;")
		fString = Replace(fString, CHR(39), "&#39;")
		fString = Replace(fString, CHR(13), "")
		fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
		fString = Replace(fString, CHR(10), "<br /> ")
	
		'fString=ChkBadWords(fString)
		HTMLEncode = fString
	End if
End function

%>