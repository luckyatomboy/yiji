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
<title><%=dianming%> - �����ѯ</title>
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
if fla1="0" and fla2="0" and fla3="0" and fla4="0" then
%>
<br><center><img src="../images/note.gif" align="absmiddle">&nbsp;<font color="#FF0000">�㲻�߱���Ȩ�ޣ��������Ա��ϵ��</font></center>
<%  
  response.end
end if
%>

<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../images/r_1.gif" alt="" /></td>
<td width="100%" background="../images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td>&nbsp;�����ѯ</td>
	  <td align="right">&nbsp;</td>
    </tr>
  </table>
</td>
<td><img src="../images/r_2.gif" alt="" /></td>
</tr>

<tr>
<td></td>
<td>
<!--startprint-->
<table align="center" cellpadding="4" cellspacing="1" class="toptable grid" border="1">

  <tr align="center">
  	<td class="category" width="120" height="30">�ֻ���ͬ���</td>
  	<td class="category" width="80" height="30">��˾����</td>
  	<td class="category" width="80" height="30">�ͻ�</td>
  	<td class="category" width="80" height="30">����</td>
  	<td class="category" width="80" height="30">����(��)</td>
  </tr>
  <%
    sql="select * from salescontract where category='A' and refshipment="&request("refshipment")&" and refitem="&request("refitem")
    set rs_contract=conn.execute(sql)    
    if not rs_contract.eof then
    do while rs_contract.eof=false    
  %>
  <tr align="center">
  	<td align="center" height="30"><%=rs_contract("contractnum")%></td>
    <td align="center"><%=rs_contract("owncompany")%></td>	  
    <td align="center"><%=rs_contract("customer")%></td>
    <td align="center"><%=rs_contract("quantity")%></td>
    <td align="center"><%=rs_contract("weight")%></td>
  </tr>
  <%
  	'end if
      rs_contract.movenext
    loop
    else
  %>
    <tr align="center" onMouseOver="this.className='highlight'" onMouseOut="this.className=''">
      <td colspan="12" height="25" align="center" style="color:red"><b>û���ҵ���¼</b></td>
    </tr>
  <%
    end if
  %>

</table>

</td>
<td></td>
</tr>

</table>

</body>
</html>