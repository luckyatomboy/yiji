<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
if session("redboy_username")="" then
%>
<script language="javascript">
top.location.href="../index.asp"
</script>
<%  
  response.end
end if

'为了使用ajax，不能include conn2.asp!!!'
set conn=server.createobject("adodb.connection")
mypath=server.mappath("../data/red#jxc.mdb")
conn.open "driver={microsoft access driver (*.mdb)};dbq="&mypath&";uid=;pwd=123;"

sql="select * from stockdocument where refshipment="&request("refshipment")&" and refitem="&request("refitem")
set rs_remain=conn.execute(sql)	

if rs_remain.eof=false then
	response.write(rs_remain("remainqty"))
else
	response.write("e")
end if

response.end()
%>