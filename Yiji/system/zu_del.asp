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
id=replace(request("id")," ","")
id=split(id,",")
for i=0 to UBound(id)
  sql="delete from login where id<>1 and id_zu="&id(i)
  conn.execute(sql)
  sql="delete from zu_login where id="&id(i)
  conn.execute(sql)
next
response.redirect "zu.asp"
%>

