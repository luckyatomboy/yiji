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

'删除逻辑锁'
sql="delete from locktable where tablename='"&request("tablename")&"' and combinedkey='"&request("combinedkey")&"'"
conn.execute(sql)

%>
<script language="javascript">
window.close();
</script>