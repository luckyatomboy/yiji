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
<title>删除订货合同</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style/style.css" rel="stylesheet" type="text/css">
</HEAD>

<%

'如果状态不是新建，则不许删除
if request("status")<>"01" then
%>
<script language="javascript">
	alert("该订货合同已被处理，不能删除！")
	window.close();
</script>
<%

dim lv_deny
lv_deny=1
end if

if lv_deny <> 1 then
'删除订货合同'
sql="delete from salescontract where contractnum="&request("contractnum")
conn.execute(sql)

'更新入库单中剩余库存'
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


'删除逻辑锁'
sql="delete from locktable where tablename='salescontract' and combinedkey='"&request("contractnum")&"'"
conn.execute(sql)
'response.redirect "shipment.asp"
%>

<script language="javascript">
	window.close();
</script>

</html>