<!--#include file="admin_yz.asp"-->
<!--#include file="conn.asp"-->
<!--#include file="lock.asp"-->
<%
conn.execute("insert into lockip (ip) values('"&request("ip")&"')")
call connclose
call errmsg("���û�IP�Ѿ��ɹ�����")
%>