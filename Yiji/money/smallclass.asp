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
<title><%=dianming%> - С�����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style/style.css" rel="stylesheet" type="text/css">
<style>
body {
	background-color:#FFFFFF;
}
</style>
<script>
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')       e.checked = form.chkall.checked; 
   }
  }
</script>
</HEAD>

<BODY>
<%
if fla67="0" and session("redboy_id")<>"1" then
%>
<br><center><img src="../images/note.gif" align="absmiddle">&nbsp;<font color="#FF0000">�㲻�߱���Ȩ�ޣ��������Ա��ϵ��</font></center>
<% 
  response.end
end if

nowtype=request("type")
nowbigclass=request("bigclass")
%>
<table width="100%" border="0" cellpadding="0" cellspacing="2" align="center">
<form name="form2">
  <tr> 
    <td width="25%" height="30">&nbsp;</td>
	<td width="75%" align="right">
	  ���ͣ�
	  <select name="type" onChange="form2.submit()">
	    <option value=""></option>
		<option value="0"<%if nowtype="0" then%> selected="selected"<%end if%>>����</option>
		<option value="1"<%if nowtype="1" then%> selected="selected"<%end if%>>֧��</option>
      </select>
    <%
	if nowtype="" then
	  sql="select * from money_bigclass order by type"
	else
	  sql="select * from money_bigclass where type="&nowtype&" order by id"
	end if
	set rs_bigclass=conn.execute(sql)
	%>
	  ��ѡ����ࣺ
	  <select name="bigclass" onChange="form2.submit()">
        <option value="">��ѡ�����</option>
        <%
	do while rs_bigclass.eof=false
	%>
        <option value="<%=rs_bigclass("id")%>"<%if trim(cstr(rs_bigclass("id")))=nowbigclass then%> selected="selected"<%end if%>><%=rs_bigclass("bigclass")%></option>
        <%
	  rs_bigclass.movenext
	loop
	%>
      </select>&nbsp;</td>
  </tr>
</form>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../images/r_1.gif" alt="" /></td>
<td width="100%" background="../images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td>&nbsp;�޸� / ɾ��С��</td>
	  <td align="right"><%if fla66="1" or session("redboy_id")="1" then%><input type="button" value="���С��" onClick="window.location.href='smallclass_add.asp'" class="button"><%end if%>&nbsp;</td>
    </tr>
  </table>
</td>
<td><img src="../images/r_2.gif" alt="" /></td>
</tr>
<tr>
<td></td>
<td>
<form name="form1" action="smallclass_del.asp">
<table align="center" cellpadding="4" cellspacing="1" class="toptable grid" border="1">
  <tr align="center">
    <td height="30" class="category">����</td>
	<td class="category">��������</td>
    <td class="category">С������</td>
	<%if fla68="1" or session("redboy_id")="1" then%>
    <td width="8%" class="category">�޸�</td>
	<%end if%>
	<%if fla69="1" or session("redboy_id")="1" then%>
    <td width="5%" class="category">ɾ��</td>
	<%end if%>
  </tr>
  <%
  if nowbigclass="" then
    if nowtype="" then
      sql="select * from money_smallclass order by id_bigclass,id"
	else
	  sql="select * from money_smallclass where id_bigclass in (select id from money_bigclass where type="&nowtype&") order by id_bigclass,id"
	end if  
  else
    sql="select * from money_smallclass where id_bigclass="&nowbigclass&" order by id"
  end if	
  set rs_smallclass=conn.execute(sql)
  do while rs_smallclass.eof=false
    sql="select * from money_bigclass where id="&rs_smallclass("id_bigclass")
	set rs_bigclass=conn.execute(sql)
  %>
  <tr onMouseOver="this.className='highlight'" onMouseOut="this.className=''">
    <td align="center" height="25"><%if rs_bigclass("type")=0 then%>����<%else%>֧��<%end if%></td>
    <td align="center"><%=rs_bigclass("bigclass")%></td>
    <td align="center"><%=rs_smallclass("smallclass")%></td>
    <%if fla68="1" or session("redboy_id")="1" then%>
	<td align="center"><a href="smallclass_modi.asp?id=<%=rs_smallclass("id")%>"><img src="../images/res.gif" border="0" hspace="2" align="absmiddle">�޸�</a></td>
	<%end if%>
    <%if fla69="1" or session("redboy_id")="1" then%>
	<td align="center"><input type="checkbox" name="ID" value="<%=rs_smallclass("id")%>" style="border:0"></td>
	<%end if%>
  </tr>
  <%
    rs_smallclass.movenext
  loop
  %>
  <tr>
    <td colspan="5" height="30" class="category">
	<table cellpadding=0 cellspacing=0 width="100%">
	<tr>
	<td align="left" style="color:#FF0000;">&nbsp;</td>  
	<td align="right">
    <%if fla69="1" or session("redboy_id")="1" then%>
	<input name="chkall" type="checkbox" id="chkall" value="select" onClick="CheckAll(this.form)" style="border:0"> ȫѡ
	<input type="submit" value="ɾ ��" class="button" onClick="return confirm('��ͬʱɾ�����ڴ�С������������¼�����������أ�����\n\nȷ��Ҫɾ����С����')">
	<%end if%>
  </td>
  </tr></table></td></tr>   
</table>
</form>
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