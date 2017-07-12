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

sql="delete from customercontact where customername='"&request("customername")&"' and contact='"&request("contact")&"'"
conn.execute(sql)

response.redirect "master.asp"
%>

