<%
if session("redboy_username")="" then
%>
<script language="javascript">
top.location.href="../index.asp"
</script>
<%  
  response.end
end if
%>
<!-- #include file="../conn2.asp" -->

<html>
<head>
<title>ɾ��������ͬ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style/style.css" rel="stylesheet" type="text/css">
</HEAD>

<%

'���״̬�����½�������ɾ��
if request("status")<>"01" then
%>
<script language="javascript">
	alert("�ö�����ͬ�ѱ���������ɾ����")
	window.close();
</script>
<%

dim lv_deny
lv_deny=1
end if

if lv_deny <> 1 then
'ɾ��������ͬ'
sql="delete from salescontract where contractnum="&request("contractnum")
conn.execute(sql)

'������ⵥ��ʣ����'
sql="select * from stockdocument where refshipment="&request("refshipment")&" and refitem="&request("refitem")
set rs=server.createobject("ADODB.RecordSet")
rs.open sql,conn,1,3

if rs.eof=false then
	nowremainqty=rs("remainqty")+request("quantity")
	rs("remainqty")=nowremainqty
	rs.update
	rs.close
end if


end if


'ɾ���߼���'
sql="delete from locktable where tablename='salescontract' and combinedkey='"&request("contractnum")&"'"
conn.execute(sql)
'response.redirect "shipment.asp"
%>

<script language="javascript">
	window.close();
</script>

</html>