<!--#include file="dbconn.asp"-->
<%
'response.write("billy test")
set conn=server.createobject("adodb.connection")
mypath=server.mappath(dbname)
conn.open "driver={microsoft access driver (*.mdb)};dbq="&mypath&";uid=;pwd=123;"

function connClose
Conn.close
Set conn = Nothing
End Function
function errmsg(message)
session("err")=message
response.redirect"../gbook/err.asp"
End Function

set conn=server.createobject("ADODB.CONNECTION")
conn.open "driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath(dbname)&";uid=;pwd=123;"
set conn1=server.createobject("adodb.connection")
mypath=server.mappath(dbname)
conn1.open "driver={microsoft access driver (*.mdb)};dbq="&mypath&";uid=;pwd=123;"

%>                                                                                                                                                                                                                                                                
