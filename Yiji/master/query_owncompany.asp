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
<title><%=dianming%> - ��˾��˾��ѯ</title>
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
if fla13="0" and session("redboy_id")<>"1" then
%>
<br><center><img src="../images/note.gif" align="absmiddle">&nbsp;<font color="#FF0000">�㲻�߱���Ȩ�ޣ��������Ա��ϵ��</font></center>
<%  
  response.end
end if

'ȡ�������ؼ���  
nowkeyword=request("keyword") 
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" align="center">
<form name="form2">
  <input type="hidden" name="form" value="<%=request("form")%>">
  <input type="hidden" name="field" value="<%=request("field")%>">
  <input type="hidden" name="field2" value="<%=request("field2")%>">
  <input type="hidden" name="span1" value="<%=request("span1")%>">
  <tr> 
<!--    <td width="50" height="21"></td> -->
	<td align="right">
	  ������
	  <input type="text" name="keyword" size="20" value="<%=nowkeyword%>">
	  <input type="submit" value=" ��ѯ " class="button">&nbsp;
	</td>
  </tr>
</form>  
</table>
<%
  if nowkeyword<>"" then
  	sql="select * from owncompany where fullname like '%"&nowkeyword&"%' order by fullname ASC"
	else
		sql="select * from owncompany order by fullname ASC"
  end if
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../images/r_1.gif" alt="" /></td>
<td width="100%" background="../images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td>&nbsp;��˾��˾��ѯ</td>
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
<form name="form1">
  <input type="hidden" name="form" value="<%=request("form")%>">
  <input type="hidden" name="keyword" value="<%=nowkeyword%>">
  <tr align="center">
	<td class="category" width="100" height="30">��˾����</td>
  <td class="category" width="200" height="30">��˾ȫ��</td>  
	<td class="category" width="300" height="30">��˾��ַ</td>
	<td class="category" width="120" height="30">�绰</td>
	<td class="category" width="120" height="30">����</td>
	<td class="category" width="120" height="30">�����ʼ�</td>
  <td class="category" width="150" height="30">�����˺�</td>
  <td class="category" width="200" height="30">������</td>
	<%if fla14="1" or session("redboy_id")="1" then%>
    <td class="category" width="60">�޸�</td>
	<%end if%>
  </tr>
  <%
  set rs_company =server.createobject("ADODB.RecordSet")	
  rs_company.open sql,conn,1,3
  if not rs_company.eof then
  do while rs_company.eof=false
  %>
  <tr align="center">
  <td align="center"><%=rs_company("company")%></td>
  <td align="center"><%=rs_company("fullname")%></td>    	  
  <td align="center"><%=rs_company("address")%></td>
  <td align="center"><%=rs_company("tel")%></td>
  <td align="center"><%=rs_company("fax")%></td>
  <td align="center"><%=rs_company("email")%></td>
  <td align="center"><%=rs_company("bankaccount")%></td>
  <td align="center"><%=rs_company("bankname")%></td>
	<%if fla14="1" or session("redboy_id")="1" then%>
    <td align="center">
    	<a href="change_owncompany.asp?form=<%=request("form")%>&company=<%=rs_company("company")%>&keyword=<%=nowkeyword%>"><img src="../images/res.gif" border="0" hspace="2" align="absmiddle">�޸�</a></td>
	<%end if%>
  </tr>
  <%
  	'end if
    rs_company.movenext
  loop
  else
  %>
  <tr align="center" onMouseOver="this.className='highlight'" onMouseOut="this.className=''">
    <td colspan="12" height="25" align="center" style="color:red"><b>û���ҵ���¼</b></td>
  </tr>
  <%
  end if
  %>
  <%
  if rs_company.recordcount>0 then 
  %> 
</table></td></tr>
<%end if%> 
</form>   
</table>
<!--endprint-->
</td>
<td></td>
</tr>

</table>

</body>
</html>