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
<title><%=dianming%> - 员工管理</title>
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
if fla26="0" and session("redboy_id")<>"1" then
%>
<br><center><img src="../images/note.gif" align="absmiddle">&nbsp;<font color="#FF0000">你不具备此权限，请与管理员联系！</font></center>
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
      <td>&nbsp;修改 / 删除员工</td>
	  <td align="right"><%if fla23="1" or session("redboy_id")="1" then%><input type="button" value="添加员工" onClick="window.location.href='user_add.asp'" class="button"><%end if%>&nbsp;</td>
    </tr>
  </table>
</td>
<td><img src="../images/r_2.gif" alt="" /></td>
</tr>
<tr>
<td></td>
<td>
<form name="form1" action="user_del.asp">
<table align="center" cellpadding="4" cellspacing="1" class="toptable grid" border="1">
  <tr align="center">
    <td height="30" class="category">员工编号</td>
		<td class="category">员工姓名</td>
		<td class="category">职务</td>
		<td class="category">性别</td>
    <td class="category">电话</td>
    <%if fla24="1" or session("redboy_id")="1" then%>
	<td width="8%" class="category">修改</td>
	<%end if%>
	<%if fla25="1" or session("redboy_id")="1" then%>
    <td width="5%" class="category">删除</td>
	<%end if%>
  </tr>
  <%
  sql="select * from login order by bianhao"
  set rs_login=conn.execute(sql)
  do while rs_login.eof=false
  %>
  <tr onMouseOver="this.className='highlight'" onMouseOut="this.className=''" onDblClick="javascript:var win=window.open('user_show.asp?id=<%=rs_login("id")%>','员工详细信息','width=853,height=470,top=176,left=161,toolbar=no,location=no,directories=no,status=no,menubar=no,resizable=no,scrollbars=yes'); win.focus()">
    <td align="center" height="25"><%=rs_login("bianhao")%></td>
		<td align="center"><%=rs_login("username")%></td>
		<td align="center"><%=rs_login("position")%></td>
    <td align="center"><%=rs_login("xinbie")%></td>
		<td align="center"><%=rs_login("tel")%></td>
	<%if fla24="1" or session("redboy_id")="1" then%>
    <td align="center"><a href="user_modi.asp?id=<%=rs_login("id")%>"><img src="../images/res.gif" border="0" hspace="2" align="absmiddle">修改</a></td>
	<%end if%>
	<%if fla25="1" or session("redboy_id")="1" then%>
    <td align="center"><%if rs_login("id")="1" then%>--<%else%><input type="checkbox" name="ID" value="<%=rs_login("id")%>" style="border:0"><%end if%></td>
	<%end if%>
  </tr>
  <%
    rs_login.movenext
  loop
  %>
  <tr>
    <td colspan="7" height="30" class="category">
	<table cellpadding=0 cellspacing=0 width="100%">
	<tr>
	<td align="left" style="color:#FF0000;">双击每行可查看员工详细资料</td>  
	<td align="right">
    <%if fla25="1" or session("redboy_id")="1" then%>
	<input name="chkall" type="checkbox" id="chkall" value="select" onClick="CheckAll(this.form)" style="border:0"> 全选
	<input type="submit" value="删 除" class="button" onClick="return confirm('此操作无法恢复！！！请慎重！！！\n\n确定要删除员工吗？')"> 
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