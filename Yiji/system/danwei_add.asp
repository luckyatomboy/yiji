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
<title><%=dianming%> - 计量单位设置</title>
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
if fla55="0" and session("redboy_id")<>"1" then
%>
<br><center><img src="../images/note.gif" align="absmiddle">&nbsp;<font color="#FF0000">你不具备此权限，请与管理员联系！</font></center>
<%  
  response.end
end if
%>
<%if request("hid1")="" then%>
<script language="javascript">
function check()
{
if (document.form1.danwei.value=="")
{
alert("有*号的必须填写！");
return false;
}
}
</script>
<form name="form1">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../images/r_1.gif" alt="" /></td>
<td width="100%" background="../images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td>&nbsp;添加计量单位(带*号的为必填项)</td>
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
    <td width="25%" height="30" align="right">计量单位名称：</td>
    <td width="75%" class="category">
        <input type="text" name="danwei" style="width:200px">
        &nbsp;<font color="#ff0000">*</font></td>
  </tr>
  <tr>
    <td height="30">&nbsp;</td>
    <td class="category">
        <input name="submit" type="submit" onClick="return check()" value=" 确 认 " class="button">
      &nbsp;&nbsp;&nbsp;&nbsp;
        <input type="hidden" name="hid1" value="ok">
        <input name="reset" type="reset" value=" 重新填写 " class="button">
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
nowdanwei=request("danwei")
sql="select * from danwei where danwei='"&nowdanwei&"'"
set rs=conn.execute(sql)
if rs.eof=false then
%>
<script language="javascript">
alert("您输入的计量单位名已经存在！")
window.history.go(-1)
</script> 
<%
  response.end
end if
sql="insert into danwei(danwei) values('"&nowdanwei&"')"
conn.execute(sql)
%>
<script language="javascript">
alert("计量单位添加成功！")
window.location.href="danwei.asp"
</script> 
<%
end if
%>
</body>
</html>