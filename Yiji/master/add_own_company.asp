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
<title><%=dianming%> - ��ӱ�˾��˾��Ϣ</title>
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
if (document.form1.company.value==""||document.form1.userid.value=="")
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
      <td>&nbsp;��ӹ�˾��Ϣ(��*�ŵ�Ϊ������)</td>
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
        <td width="25%" height="30" align="right">��˾��ƣ�</td>
        <td width="75%" class="category"><input type="text" name="company" style="width:100px"> 
        	&nbsp;<font color="#ff0000">*</font></td>
      </tr>    
      <tr>
        <td align="right" height="30">��˾ȫ�ƣ�</td>
        <td class="category"><input type="text" name="fullname" style="width:250px"></td>
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
        <td class="category"><input type="text" name="email" style="width:250px"></td>
      </tr> 	
      <tr>
        <td align="right" height="30">�����˻�����1��</td>
        <td class="category">
          <input type="text" name="bankaccountname" style="width:200px">
          &nbsp;&nbsp;&nbsp;&nbsp;�����˺�1
          <input type="text" name="bankaccount" style="width:200px">
          &nbsp;&nbsp;&nbsp;&nbsp;����������1
          <input type="text" name="bankname" style="width:200px">
        </td>
      </tr>            
      <tr>
        <td align="right" height="30">�����˻�����2��</td>
        <td class="category">
          <input type="text" name="bankaccountname2" style="width:200px">
          &nbsp;&nbsp;&nbsp;&nbsp;�����˺�2
          <input type="text" name="bankaccount2" style="width:200px">
          &nbsp;&nbsp;&nbsp;&nbsp;����������2
          <input type="text" name="bankname2" style="width:200px">
        </td>
      </tr>          
      <tr>
        <td align="right" height="30">��ע��</td>
        <td class="category"><textarea name="memo" cols="70" rows="4"></textarea></td>
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
nowcompany=request("company")
nowfullname=request("fullname")
nowtel=request("tel")
nowaddress=request("address")
nowfax=request("fax")
nowemail=request("email")
nowaccount=request("bankaccount")
nowbankname=request("bankname")
nowaccountname=request("bankaccountname")
nowaccount2=request("bankaccount2")
nowbankname2=request("bankname2")
nowaccountname2=request("bankaccountname2")
nowmemo=request("memo")
sql="select * from Owncompany where Company='"&nowcompany&"'"
set rs=conn.execute(sql)
if rs.eof=false then
%>
<script language="javascript">
alert("������Ĵ���˾�Ѿ����ڣ����������룡")
window.history.go(-1)
</script> 
<%
  response.end
end if

sql="insert into Owncompany(Company,Fullname,Address,Tel,Fax,Email,bankaccount,bankname,bankaccountname,bankaccount2,bankname2,bankaccountname2,Memo,CreateDate,Creator) values('"&nowcompany&"','"&nowfullname&"','"&nowaddress&"','"&nowtel&"','"&nowfax&"','"&nowemail&"','"&nowaccount&"','"&nowbankname&"','"&nowaccountname&"','"&nowaccount2&"','"&nowbankname2&"','"&nowaccountname2&"','"&nowmemo&"',#"&now()&"#,'"&session("redboy_username")&"')"
conn.execute(sql)
%>
<script language="javascript">
alert("��˾��˾��ӳɹ���")
window.location.href="master.asp"
</script> 
<%
end if
%>
</body>
</html>