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
<%

sql="select from salescontract where contractnum='"&request("contractnum")&"'"
rs=conn.execute(sql)
'如果状态不是新建，则不许删除
if rs("status")<>"01" then
%>
<script language="javascript">
	alert("订货合同已经被处理，不能删除！")
	window.history.go(-1);
</script>
<%
end if

'删除订货合同'
sql="delete from salescontract where contractnum='"&request("contractnum")&"'"
conn.execute(sql)

'删除逻辑锁'
sql="delete from locktable where tablename='salescontract' and combinedkey='"&request("contractnum")&"'"
conn.execute(sql)

response.redirect "shipment.asp"
%>

