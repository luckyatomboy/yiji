<%response.expires=0%>
<!--#include file="conn.asp"-->
<!-- #include file="../const.asp" --> 
<%
if fla92="0" and session("redboy_id")<>"1" then
%>
<br><center><img src="../images/note.gif" align="absmiddle">&nbsp;<font color="#FF0000">�㲻�߱���Ȩ�ޣ��������Ա��ϵ��</font></center>
<%  
  response.end
end if
%>
<html>
<head>
<title>��ҵͨѸ¼</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<LINK href="style.css" type=text/css rel=stylesheet>	

<frameset cols="150,*" frameborder="YES" border="1" framespacing="1" bordercolor="#3366FF"> 
  <frame name="leftFrame" scrolling="auto" noresize src="Left.asp">
  <frame name="mainFrame" src="">
</frameset>
<noframes><body bgcolor="#FFFFFF" text="#000000">
<center><font color="red" size="+1">�Բ��������������֧�ֿ�ܣ�</font></center>
</body></noframes>
</html>
