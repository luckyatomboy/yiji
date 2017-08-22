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
technum=replace(request("technum")," ","")
technum=split(technum,",")
for i=0 to UBound(technum)
  sql="delete from locktable where technum="&technum(i)
  conn.execute(sql)
next
response.redirect "user.asp"
%>

