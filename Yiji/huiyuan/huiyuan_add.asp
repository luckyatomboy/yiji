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
<title><%=dianming%> - ��ӿͻ�</title>
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
if fla5="0"then
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
if (document.form1.username.value==""||document.form1.userid.value=="")
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
      <td>&nbsp;��ӿͻ�(��*�ŵ�Ϊ������)</td>
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
        <td width="25%" height="30" align="right">�ͻ����룺</td>
        <td width="75%" class="category"><input type="text" name="userid" style="width:60px"> 
        	&nbsp;<font color="#ff0000">*</font></td>
      </tr>
      <tr>
        <td align="right">�ͻ����ƣ�</td>
        <td class="category"><input type="text" name="username" style="width:200px"></td>
      </tr>      
      <tr>
        <td align="right" height="30">��˾��ַ��</td>
        <td class="category"><input type="text" name="address" style="width:300px"></td>
      </tr> 	  
      <tr>
        <td align="right" height="30">��ϵ�绰��</td>
        <td class="category"><input type="text" name="tel" style="width:200px"></td>
      </tr>	 
      <tr>
        <td align="right" height="30">���棺</td>
        <td class="category"><input type="text" name="fax" style="width:200px"></td>
      </tr>	       
	    <tr>
        <td align="right" height="30">Email��</td>
        <td class="category"><input type="text" name="email" style="width:200px"></td>
      </tr> 
      <tr>
        <td align="right" height="30">��ϵ�ˣ�</td>
        <td class="category"><input type="text" name="partner" style="width:200px"></td>
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
nowuserid=request("userid")
nowusername=request("username")
nowtel=request("tel")
nowaddress=request("address")
nowfax=request("fax")
nowemail=request("email")
nowpartner=request("partner")
nowbeizhu=request("beizhu")
sql="select * from Customer where CustomerID="&nowuserid&""
set rs=conn.execute(sql)
if rs.eof=false then
%>
<script language="javascript">
alert("������Ŀͻ�����Ѿ����ڣ����������룡")
window.history.go(-1)
</script> 
<%
  response.end
end if

sql="insert into Customer(CustomerID,CustomerName,Address,Tel,Fax,Email,Partner,Memo,CreateDate,Creator) values("&nowuserid&",'"&nowusername&"','"&nowaddress&"','"&nowtel&"','"&nowfax&"','"&nowemail&"','"&nowpartner&"','"&nowbeizhu&"',#"&now()&"#,'"&session("redboy_username")&"')"
conn.execute(sql)
%>
<script language="javascript">
alert("�ͻ���ӳɹ���")
window.location.href="huiyuan.asp"
</script> 
<%
end if
%>
</body>
</html>