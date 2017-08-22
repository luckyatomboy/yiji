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
<title><%=dianming%> - 解锁记录</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style/style.css" rel="stylesheet" type="text/css">
<style>
body {
	background-color:#FFFFFF;
}
</style>
</HEAD>

<BODY>
<script>
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')       e.checked = form.chkall.checked; 
   }
  }
</script>

<%
'取得搜索关键字  
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
	  搜索：
	  <input type="text" name="keyword" placeholder="请输入用户名" size="20" value="<%=nowkeyword%>">
	  <input type="submit" value=" 查询 " class="button">&nbsp;
	</td>
  </tr>
</form>  
</table>
<%
  if nowkeyword<>"" then
  	sql="select * from locktable where username like '%"&nowkeyword&"%'"
	else
		sql="select * from locktable order by username ASC"
  end if
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../images/r_1.gif" alt="" /></td>
<td width="100%" background="../images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td>&nbsp;记录锁查询</td>
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
<form name="form1" action="del_lock.asp">
  <input type="hidden" name="form" value="<%=request("form")%>">
  <input type="hidden" name="keyword" value="<%=nowkeyword%>">
  <tr align="center">
	<td class="category" width="180" height="30">用户名</td>
	<td class="category" width="380" height="30">表名</td>
	<td class="category" width="120" height="30">主键</td>
	<td class="category" width="120" height="30">锁定时间</td>
	<%if fla14="1" or session("redboy_id")="1" then%>
    <td class="category" width="60">删除</td>
	<%end if%>
  </tr>
  <%
  set rs_lock =server.createobject("ADODB.RecordSet")	
  rs_lock.open sql,conn,1,3
  if not rs_lock.eof then
  do while rs_lock.eof=false
  %>
  <tr align="center">
  <td align="center"><%=rs_lock("username")%></td>	  
  <td align="center"><%=rs_lock("tablename")%></td>
  <td align="center"><%=rs_lock("combinedkey")%></td>
  <td align="center"><%=rs_lock("locktime")%></td>
	<%if fla14="1" or session("redboy_id")="1" then%>
    	<td align="center"><input type="checkbox" name="technum" value="<%=rs_lock("technum")%>" style="border:0"></td>
	<%end if%>
  </tr>
  <%
  	'end if
    rs_lock.movenext
  loop
  %>
  <tr>
      <td colspan="7" height="30" class="category">
    <table cellpadding=0 cellspacing=0 width="100%">
    <tr>
    <td align="left"></td>  
    <td align="right">
      <%if fla25="1" or session("redboy_id")="1" then%>
    <input name="chkall" type="checkbox" id="chkall" value="select" onClick="CheckAll(this.form)" style="border:0"> 全选
    <input type="submit" value="删 除" class="button" onClick="return confirm('是否确定删除该记录锁？')"> 
    <%end if%>
    </td>
    </tr></table></td></tr>    
  <%
  else
  %>
  <tr align="center" onMouseOver="this.className='highlight'" onMouseOut="this.className=''">
    <td colspan="12" height="25" align="center" style="color:red"><b>没有找到记录</b></td>
  </tr>
  <%
  end if
  %>

  <%
  if rs_lock.recordcount>0 then 
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