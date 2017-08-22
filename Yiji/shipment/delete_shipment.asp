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

sql="select from shipment where shipmentnum="&request("shipmentnum")
rs=conn.execute(sql)

'查询是否有订货合同。如果有，则不许删除船期表'
sql="select from salescontract where refshipment="&request("shipmentnum")
rs_contract=conn.execute(sql)
if rs_contract.eof=false then
%>
<script language="javascript">
	alert("订货合同已经存在，不能删除船期表！")
	window.history.go(-1);
</script>
<%
end if

'查询是否有入库单。如果有，则不许删除船期表'
sql="select from stocknumber where refshipment="&request("shipmentnum")
rs_stock=conn.execute(sql)
if rs_stock.eof=false then
%>
<script language="javascript">
	alert("入库单已经存在，不能删除船期表！")
	window.history.go(-1);
</script>
<%
end if

'删除船期表'
sql="delete from shipment where shipmentnum="&request("shipmentnum")
conn.execute(sql)
sql="delete from shipmentitem where shipmentnum="&request("shipmentnum")
conn.execute(sql)

'删除逻辑锁'
sql="delete from locktable where tablename='shipment' and combinedkey='"&request("shipmentnum")&"'"
conn.execute(sql)

response.redirect "shipment.asp"
%>

