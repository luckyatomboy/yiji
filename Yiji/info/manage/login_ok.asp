<%
adminname=request.form("adminname")
adminpass=request.form("adminpass")
admincode=request.form("admincode")

if adminname="" or adminpass="" then
	response.write("<script language='javascript'>alert('�û��������벻��Ϊ�գ������룡');history.back();</script>")
	response.end
end if
if admincode="" or admincode<>cstr(session("admincode")) then
	response.write("<script language='javascript'>alert('������������������룡');history.back();</script>")
	response.end
end if
%>
<!--#include file="../include/conn.asp"-->
<%
sqlstr="select adminid from tbl_admin where adminname='"&adminname&"' and adminpass='"&adminpass&"'"
set rs=conn.execute(sqlstr)
if rs.eof then
	response.write("<script language='javascript'>alert('�û��������벻��ȷ�������룡');history.back();</script>")
	response.end
end if
session("adminname")=adminname

response.redirect "admin_index.asp"
%>