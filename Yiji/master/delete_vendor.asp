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

sql="delete from vendor where vendorname='"&request("vendorname")&"'"
conn.execute(sql)

'删除逻辑锁'
sql="delete from locktable where tablename='vendor' and combinedkey='"&request("vendorname")&"'"
conn.execute(sql)

response.redirect "master.asp"
%>

