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

sql="delete from piwen where company='"&request("company")&"'"
conn.execute(sql)

response.redirect "master.asp"
%>

