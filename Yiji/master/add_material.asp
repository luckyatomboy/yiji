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
<title><%=dianming%> - ���Ʒ��</title>
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
if fla35="0" and session("redboy_id")<>"1" then
%>
<br><center><img src="../images/note.gif" align="absmiddle">&nbsp;<font color="#FF0000">�㲻�߱���Ȩ�ޣ��������Ա��ϵ��</font></center>
<%  
  response.end
end if
%>

<%if request("hid1")="" then%> 
<script language="javascript">
function check()
{
if(document.form1.material.value=="")
{
alert("��*�ŵı�����д��");
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
      <td>&nbsp;���Ʒ��(��*�ŵ�Ϊ������)</td>
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
        <td align="right">Ʒ����</td>
        <td class="category"><input type="text" name="material" style="width:150px"> 
        	&nbsp;<font color="#ff0000">*</font></td>
      </tr>             
      <tr>
        <td align="right" height="30">��ע��</td>
        <td class="category"><textarea name="beizhu" cols="70" rows="4"></textarea></td>
      </tr>	  
      <tr>
	    <td height="30">&nbsp;</td>
        <td class="category">
		  <input type="submit" value=" ȷ����� " onClick="return check()" class="button">&nbsp;&nbsp;&nbsp;&nbsp;
		  <input type="hidden" name="hid1" value="ok">
		  <input type="reset" value=" ������д " class="button">		</td>
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
nowmaterial=request("material")
nowdes=request("beizhu")
sql="select * from material where MaterialName='"&nowmaterial&"'"
set rs=conn.execute(sql)
if rs.eof=false then
%>
<script language="javascript">
alert("�������Ʒ���Ѿ����ڣ����������룡")
window.history.go(-1)
</script> 
<%
  response.end
end if

sql="insert into Material(MaterialName,Description,CreateDate,Creator) values('"&nowmaterial&"','"&nowdes&"',#"&now()&"#,'"&session("redboy_username")&"')"
conn.execute(sql)
%>
<script language="javascript">
alert("Ʒ����ӳɹ���")
window.location.href="master.asp"
</script> 
<%
end if
%>
</body>
</html>