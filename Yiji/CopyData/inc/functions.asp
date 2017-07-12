<%
'===============================================================
'�������ƣ��ַ��������ַ�ת����ʹ�ܹ��������ݿ��
'��    ����
'----------str ---- ��ת�����ַ���
'�� �� ֵ��ת��������ַ���
function FormatString(str)
	str = replace(str,chr(39),"��")	'������
	str = replace(str,chr(34),"��")	'˫����
	str = replace(str,chr(44),"��")	'����
	'str = replace(str,chr(32) & chr(32),"&nbsp;&nbsp;")	'˫�ո�
	str = replace(str,chr(62),"&gt;")	'���ں�
	str = replace(str,chr(60),"&lt;")	'С�ں�
	'str = replace(str,vbcrlf,"<br>")	'�س�
	'...................
	FormatString = str
end function
'===============================================================
'�������ƣ����ַ�����HTML��ʽ����
'��    ����
'----------str ---- ��ת�����ַ���
'�� �� ֵ��ת��������ַ���
function FormatStringHTML(str)
	str = replace(str,chr(32),"&nbsp;")	'˫�ո�
	str = replace(str,vbcrlf,"<br>")	'�س�����
	'...................
	FormatStringHTML = str
end function
'===============================================================
'�������ƣ������ڸ�ʽ��Ϊ��Ҫ�ĸ�ʽ
'��    ����
'----------strDate ---- ��ת�����ַ���
'�� �� ֵ��ת��������ַ���
function FormatDate(strDate,Date_Type)
	y = year(strDate)
	m = month(strDate)
	d = day(strDate)
	select case Date_Type
		case 1	'yyyy-m-d ��ʽ
			FormatDate = y & "-" & m & "-" & d
		case 2	'yyyy/m/d ��ʽ
			FormatDate = y & "/" & m & "/" & d
		case 3	'm/d/yyyy ��ʽ
			FormatDate = m & "/" & d & "/" & y
		case 4	'��-��-�� ��ʽ
			FormatDate = y & "��" & m & "��" & d & "��"
	end select
end function
'===============================================================
'�������ƣ�Ȩ����֤
'��    ����
'----------Employee_ID ---- Ա�����
'----------Application_ID ----  ���ܱ�ţ�ģ���ţ�3λ��+���ܱ�ţ�4λ�������磺2000200�е�ǰ3λ��200����ʾ����ͬ��������4λ��0200����ʾ����ͬ�����еġ���ǩ��ͬ��
'�� �� ֵ������ֵ
Function CheckAuth(Employee_ID,Application_ID)
	CheckAuth = False	'Ĭ��û�и�Ȩ��
	if Employee_ID = 1 then	'�����û���admin����Employee_ID Ϊ��1����һ����ӵ��û���ӵ������Ȩ��
		CheckAuth = true
		exit function
	end if	
	if IsNumeric(Application_ID) then	'��ȷ�Ĺ��ܱ���
		'����������ж��Ƿ���Ȩ��
		'select Employee_ID from Employee_Auth where Employee_ID = @Employee_ID and Application_ID = @Application_ID
		'sql = "sp_Employee_Check_Auth " & Employee_ID & "," & Application_ID
		sql = "select Employee_ID from Employee_Auth where Employee_ID = " & Employee_ID & " and Application_ID = " & Application_ID
		set conn = Server.CreateObject("ADODB.Connection")
		set rs = Server.CreateObject("ADODB.Recordset")
		conn.open connstr
		set rs = conn.execute(sql)
		if not rs.eof then
			CheckAuth = true
		else
			CheckAuth = false
		end if
		rs.close
		set rs = nothing
		conn.close
		set conn = nothing
	else	'����Ĺ��ܱ���
		CheckAuth = False
		exit function
	end if
End Function

'===============================================================
'�������ƣ�������־
'��    ����
'----------Rec_Content ----  ��������
'----------Command_Line ---- ִ��������
'�� �� ֵ����
Sub Op_Rec(Rec_Content,Command_Line)
	'ת�������ַ�
	Command_Line = FormatString(Command_Line)
	'insert into Employee_Op_Rec(Employee_ID,Rec_Content,Remote_Addr,Command_Line)  values(@Employee_ID,@Rec_Content,@Remote_Addr,@Command_Line)
	'sql = "sp_Employee_Op_Rec " & Session("Employee_ID") & ",'" & Rec_Content & "','" & Request.ServerVariables("REMOTE_ADDR") & "','" & Command_Line & "'"
	'response.write sql
	sql = "insert into Employee_Op_Rec(Employee_ID,Rec_Content,Remote_Addr,Command_Line)  values(" & Session("Employee_ID") & ",'" & Rec_Content & "','" & Request.ServerVariables("REMOTE_ADDR") & "','" & Command_Line & "')"
	set conn = Server.CreateObject("ADODB.Connection")
	conn.open connstr
	conn.execute(sql)
	conn.close
	set conn = nothing
End Sub

'===============================================================
'�������ƣ�������ҵ������
'��    ����
'----------Trade_Code ---- Ĭ����ҵ����
'�� �� ֵ����
sub Trade_List(Trade_Code)
	response.write "<select name=""Trade_Code"" size=1>"
	for i=0 to UBound(aryTrade)
		if CStr(Trade_Code) = CStr(i) then tag = " selected"
		response.write "<option value=""" & i & """" & tag & ">" & aryTrade(i) & "</option>"
		tag = ""
	next
	response.write "</select>"
end sub

'===============================================================
'�������ƣ����㹤����
'��    ����
'----------date_begin ---- ��ʼ����
'----------date_end ---- ��������
'�� �� ֵ��������
function workdays(date_begin , date_end)
	date_begin = CDate(date_begin)
	date_end = CDate(date_end)
	if IsDate(date_begin) and IsDate(date_end) then
		workdays = 0
		for d=0 to datediff("d" , date_begin , date_end)
			if weekday(date_begin + d) <> 1 and weekday(date_begin + d) <> 7 then
				workdays = workdays + 1
			end if
		next
	else	
		workdays = "δ֪"
	end if
end function


'===============================================================
'�������ƣ��ж��Ƿ���Ϲ淶
'��    ����
'----------Str ---- ���ж��ַ���
'----------Type ---- ���ж��ַ������ͣ�n-���֣�s-�ַ�����d-����
'�� �� ֵ����
function IsNot(InputString,StringType)
	'Ĭ�����벻�Ϸ�
	IsNot = false
	'����Ϊ��ʱ���Ϸ�
	if Trim(InputString) = "" then exit function
	'�Ϸ�����
	select case StringType
		case "n"
			if IsNumeric(InputString) then IsNot = true
		case "d"
			if IsDate(InputString) then IsNot = true
		case "s"
			IsNot = true
		case else
			IsNot = false	'����֪�����Ϊ���Ϸ�
	end select
end function
%>