<%
If Session("ManageName")="" or Session("ManagePwd")="" or Session("ManageLevel")="" Then
				response.write("<script language='javascript'>alert('��½��Ϣ���ڣ������µ�½!');self.close();</script>")
			response.end	
end if
		Session("Employee_Name") = "ϵͳ����Ա"
		Session("Employee_ID") = 1
if Trim(Session("Employee_ID")) = "" then
	response.write "<script language=""JavaScript"" type=""text/javascript"">"
	response.write "alert('��û�е�¼���߲�����ʱ�����ȵ�¼��');"
	response.write "</script>"
	response.end
end if
%>