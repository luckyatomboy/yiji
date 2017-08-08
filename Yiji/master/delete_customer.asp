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

sql="delete from customer where customername='"&request("customername")&"'"
conn.execute(sql)

'删除逻辑锁'
sql="delete from locktable where tablename='customer' and combinedkey='"&request("customername")&"'"
conn.execute(sql)

response.redirect "master.asp"
%>

