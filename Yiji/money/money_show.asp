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
<!-- #include file="../const.asp" -->
<html>
<head>
<title><%=dianming%> - ��Ʒ��ϸ��Ϣ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style/style.css" rel="stylesheet" type="text/css">
<style>
body {
	background-color:#FFFFFF;
}
</style>
</HEAD>

<BODY>
<%
if fla71="0" and session("redboy_id")<>"1" then
%>
<br><center><img src="../images/note.gif" align="absmiddle">&nbsp;<font color="#FF0000">�㲻�߱���Ȩ�ޣ��������Ա��ϵ��</font></center>
<%  
  response.end
end if
%>
<%
sql="select * from redboy_money where id="&request("id")
set rs=conn.execute(sql)
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../images/r_1.gif" alt="" /></td>
<td width="100%" background="../images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td>&nbsp;������ϸ��Ϣ</td>
	  <td align="right">&nbsp;</td>
    </tr>
  </table>
</td>
<td><img src="../images/r_2.gif" alt="" /></td>
</tr>
<tr>
<td></td>
<td>
<table align="center" cellpadding="4" cellspacing="1" class="toptable grid" border="1">
      <tr>
	    <td align="right" height="30">���ͣ�</td>
        <td class="category"><%if rs("type")=0 then%>����<%else%>֧��<%end if%></td>
      </tr>
	  <tr>
        <td align="right" height="30">�������ࣺ</td>
        <td class="category">
          <%
	      sql="select * from money_bigclass where id="&rs("id_bigclass")
	      set rs_bigclass=conn.execute(sql)
	      %>
          <%if rs_bigclass.eof=false then%><%=rs_bigclass("bigclass")%><%else%>&nbsp;<%end if%>	</td>
      </tr>  
      <tr>
        <td align="right" height="30">����С�ࣺ</td>
        <td class="category">
          <%
	      sql="select * from money_smallclass where id="&rs("id_smallclass")
	      set rs_smallclass=conn.execute(sql)
		  %>
          <%if rs_smallclass.eof=false then%><%=rs_smallclass("smallclass")%><%else%>&nbsp;<%end if%></td>
      </tr>	   
      <tr>
        <td width="25%" height="30" align="right">�� </td>
        <td width="75%" class="category"><%=rs("price")%></td>
      </tr>	  
      <tr>
        <td align="right" height="30">�����ˣ�</td>
        <td class="category">
		  <%
		  sql="select * from login where id="&rs("id_login")
		  set rs_login=conn.execute(sql)
		  %>
		  <%if rs_login.eof then%><%=rs("login")%><%else%><%=rs_login("username")%><%end if%>		
		</td>
      </tr>
	  <tr>
        <td align="right" height="30">ʱ�䣺</td>
        <td class="category">
		  <%=rs("selldate")%></td>
      </tr>	
      <tr>
        <td align="right" height="30">��ע��</td>
        <td class="category">
		  <%=replace(replace(rs("beizhu")&""," ","&nbsp;"),chr(13),"<br>")%></td>
      </tr>		  		      
</table>
</td>
<td></td>
</tr>
<tr>
<td><img src="../images/r_4.gif" alt="" /></td>
<td></td>
<td><img src="../images/r_3.gif" alt="" /></td>
</tr>
</table>
</body>
</html>