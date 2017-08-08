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

sql="select from stockdocument where stocknumber='"&request("stocknumber")&"'"
rs=conn.execute(sql)
'查询是否有现货合同。如果有，则不许删除入库单'
sql="select from salescontract where category='A' and refshipment='"&rs("refshipment")&"' and refitem='"&rs("refitem")&"'"
rs_contract=conn.execute(sql)
if rs_contract.eof=false then
%>
<script language="javascript">
	alert("现货合同已经存在，不能删除入库单！")
	window.history.go(-1);
</script>
<%
end if

'删除入库单'
sql="delete from stockdocument where stocknumber='"&request("stocknumber")&"'"
conn.execute(sql)

'删除逻辑锁'
sql="delete from locktable where tablename='stockdocument' and combinedkey='"&request("stocknumber")&"'"
conn.execute(sql)

response.redirect "shipment.asp"
%>

