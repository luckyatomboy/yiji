<!-- #include file="conn.asp" -->
<%

sql="delete from signinuser where id="&session("redboy_id")
conn.execute(sql)

session("redboy_username")=""
session("redboy_id")=""
session("redboy_zu")=""
response.Redirect "index.asp"
%>