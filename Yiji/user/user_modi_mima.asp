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
<!-- #include file="../inc/md5.asp" -->
<html>
<head>
<title><%=dianming%> - 修改密码</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style/style.css" rel="stylesheet" type="text/css">
<style>
body {
	background-color:#FFFFFF;
}
</style>
<script language="javascript">
function check()
{
if (document.form1.pwd_new.value=="")
{
alert("请输入新密码！");
document.form1.pwd_new.focus();
return false;
}
if (document.form1.pwd_new.value!=document.form1.pwd_new2.value)
{
alert("两次密码输入不一致！");
return false;
}
}
</script>
</HEAD>

<BODY>
<%if request("hid1")="" then%> 
<%
sql="select * from login where id="&session("redboy_id")
set rs=conn.execute(sql)
%>
<form name="form1">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../images/r_1.gif" alt="" /></td>
<td width="100%" background="../images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td>&nbsp;修改密码</td>
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
        <td width="25%" height="30" align="right">原密码：</td>
        <td width="75%" class="category"><input type="password" name="pwd_old" style="width:200px"></td>
      </tr>
      <tr>
        <td align="right" height="30">新密码：</td>
        <td class="category"><input type="password" name="pwd_new" style="width:200px"></td>
      </tr>
      <tr>
        <td align="right" height="30">确认新密码：</td>
        <td class="category"><input type="password" name="pwd_new2" style="width:200px"></td>
      </tr>	  	  
      <tr>
	    <td height="30">&nbsp;</td>
        <td class="category">
		  <input type="submit" value=" 确认修改 " onClick="return check()" class="button">&nbsp;&nbsp;&nbsp;&nbsp;
		  <input type="hidden" name="hid1" value="ok">
		  <input type="hidden" name="id" value="<%=rs("id")%>">
		  <input type="reset" value=" 重置 " class="button">
		  </td>
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
</form>
<%
else
pwd_old=request("pwd_old")
pwd_new=request("pwd_new")
sql="select * from login where username='"&session("redboy_username")&"' and password='"&md5(pwd_old)&"'"
set rs=conn.execute(sql)
if rs.eof then
%>
<script language="javascript">
alert("原密码不正确！")
window.history.go(-1)
</script> 
<%
  response.end
end if
sql="update login set password='"&md5(pwd_new)&"' where id="&request("id")
conn.execute(sql)
%>
<script language="javascript">
alert("密码修改成功！")
window.location.href="../desk.asp"
</script> 
<%
end if
%>
</body>
</html>