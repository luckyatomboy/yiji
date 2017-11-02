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

'if 1=2 then
sql="delete from material where material='"&request("material")&"'"
conn.execute(sql)
'end if

'删除逻辑锁'
sql="delete from locktable where tablename='material' and combinedkey='"&request("material")&"'"
conn.execute(sql)

'response.redirect "master.asp"
%>

<script language="javascript">
window.close();
</script>
