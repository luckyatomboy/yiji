<%response.expires=0%>
<!--#include file="conn.asp"-->

<html>

<head>
<meta http-equiv="expires" content="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style/style.css" rel="stylesheet" type="text/css">
<style>
body {
	background-color:#FFFFFF;
}
</style>
<title>Crackgo�칫ϵͳ</title>
</head>
<link href="../style/style.css" rel="stylesheet" type="text/css">
<style>
body {
	background-color:#FFFFFF;
}
</style>
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../images/r_1.gif" alt="" /></td>
<td width="100%" background="../images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td>&nbsp;����ͨѸ¼</td>
	  <td align="right">&nbsp;</td>
    
    </tr>
  </table>
  </table>
		</td>
	  </tr>
	</table>
<body bgcolor="#ffffff" topmargin="5" leftmargin="5">
<br>
<center><font size="+1">��ҵ������</font><br><br>
<%
set conn=dbconn("conn")
userlevel=request("userlevel")
olduserlevel=request("olduserlevel")
xh=request("xh")
id=request("id")

if not IsNumeric(xh) then
%>
<font color=red><%=xh%> ��ű��������֣�</font><br>
<%
response.end	
end if
'-----------------------------------------------
if request("submit")="����" and userlevel<>"" then
set conn=dbconn("conn")
set rs=server.createobject("adodb.recordset")
sql="select * from fenlei where leibie='" & userlevel&"'"
rs.open sql,conn,1
if not rs.eof and not rs.bof then
%>
<font color=red><%=userlevel%>�Ѿ����ڣ��뻻�����ԣ�</font><br>
<%
else
sql = "Insert Into fenlei (leibie,xh) Values('"&userlevel& "',"&xh&")"
conn.Execute sql
%>
<font color=red><%=userlevel%>���ӳɹ���</font>
<%
end if
end if
'---------------------------------------------------
if request("submit")="ɾ��" then
set conn=dbconn("conn")
sql="delete from fenlei where id="&id
conn.execute sql
%>
<font color=red><%=userlevel%>ɾ���ɹ���</font>
<%
end if
'---------------------------------------------------
if request("submit")="�޸�" and userlevel<>"" then

'�ж��Ƿ������޸ĵ�ְλ��ͬ��
set conn=dbconn("conn")
set rs=server.createobject("adodb.recordset")
sql="select * from fenlei where leibie='" & userlevel&"' and id<>"&id
rs.open sql,conn,1
if not rs.eof and not rs.bof then
%>
<font color=red><%=userlevel%>�Ѿ����ڣ��뻻�����ԣ�</font><br>
<%
else
sql = "update fenlei set leibie='" & userlevel & "',xh="&xh&" where id=" & id
conn.Execute sql
%>
<font color=red>�޸ĳɹ���</font>
<%
end if
end if
%>
<table border="1"  cellspacing="0" cellpadding="0">
<%
set conn=dbconn("conn")
set rs=server.createobject("adodb.recordset")
sql="select * from fenlei order by xh"
rs.open sql,conn,1
while not rs.eof and not rs.bof
%>
<tr>
<form method="post" action="addtype.asp">
<td>
<input type="submit" name="submit" value="ɾ��"></td><td><input type="hidden" name="olduserlevel" value="<%=rs("leibie")%>"><input type="hidden" name="id" value=<%=rs("id")%>><input type="text" name="xh"  size="6" value=<%=rs("xh")%>><input type="text" name="userlevel" value="<%=rs("leibie")%>" maxlentgh="25"></td><td><input type="submit" name="submit" value="�޸�"></td>
</form>
</tr>
<%
rs.movenext
wend
%>
</table>
<form method="post" action="addtype.asp">
<input type="text" name="xh"  size="6" ><input type="text" name="userlevel"><input type="submit" name="submit" value="����">
</form>
</center>
</body>
</html>
