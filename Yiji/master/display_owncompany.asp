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
<title><%=dianming%> - 查看本司公司</title>
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
if fla6="0" and session("redboy_id")<>"1" then
%>
<br><center><img src="../images/note.gif" align="absmiddle">&nbsp;<font color="#FF0000">你不具备此权限，请与管理员联系！</font></center>
<%  
  response.end
end if
%>

<%
sql="select * from owncompany where company='"&request("company")&"'"
set rs=conn.execute(sql)
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../images/r_1.gif" alt="" /></td>
<td width="100%" background="../images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td>&nbsp;查看本司公司资料</td>
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
  <form name="form2">
      <tr>
        <td width="25%" height="30" align="right">公司简称：</td>
        <td width="75%" class="category"><%=rs("company")%>
        	</td>
      </tr>    
      <tr>
        <td align="right" height="30">公司全称：</td>
        <td class="category"><input type="text" name="fullname" style="width:250px;border:none" readonly value="<%=rs("fullname")%>"></td>
      </tr>       
      <tr>
        <td align="right" height="30">公司地址：</td>
        <td class="category"><input type="text" name="address" style="width:300px;border:none" readonly value="<%=rs("address")%>"></td>
      </tr> 	  
      <tr>
        <td align="right" height="30">联系电话：</td>
        <td class="category"><input type="text" name="tel" style="width:200px;border:none" readonly value="<%=rs("tel")%>"></td>
      </tr>	 
      <tr>
        <td align="right" height="30">传真：</td>
        <td class="category"><input type="text" name="fax" style="width:200px;border:none" readonly value="<%=rs("fax")%>"></td>
      </tr>	       
	    <tr>
        <td align="right" height="30">Email：</td>
        <td class="category"><input type="text" name="email" style="width:250px;border:none" readonly value="<%=rs("email")%>"></td>
      </tr>    
      <tr>
        <td align="right" height="30">银行账户名称1：</td>
        <td class="category">
          <input type="text" name="bankaccountname" style="width:200px;border:none" readonly value="<%=rs("bankaccountname")%>">
          &nbsp;&nbsp;&nbsp;&nbsp;银行账号1
          <input type="text" name="bankaccount" style="width:200px;border:none" readonly value="<%=rs("bankaccount")%>">
          &nbsp;&nbsp;&nbsp;&nbsp;开户行名称1
          <input type="text" name="bankname" style="width:200px;border:none" readonly value="<%=rs("bankname")%>">
        </td>
      </tr>   
      <tr>
        <td align="right" height="30">银行账户名称2：</td>
        <td class="category">
          <input type="text" name="bankaccountname2" style="width:200px;border:none" readonly value="<%=rs("bankaccountname2")%>">
          &nbsp;&nbsp;&nbsp;&nbsp;银行账号2
          <input type="text" name="bankaccount2" style="width:200px;border:none" readonly value="<%=rs("bankaccount2")%>">
          &nbsp;&nbsp;&nbsp;&nbsp;开户行名称2
          <input type="text" name="bankname2" style="width:200px;border:none" readonly value="<%=rs("bankname2")%>">
        </td>
      </tr>                 
      <tr>
        <td align="right" height="30">备注：</td>
        <td class="category"><textarea name="memo" cols="70" rows="4" readonly><%=rs("memo")%></textarea></td>
      </tr>	   
      <tr>
        <td align="right" height="30">创建时间：</td>
        <td class="category"><%=rs("createdate")%>
      </tr>	  
      <tr>
        <td align="right" height="30">创建人：</td>
        <td class="category"><%=rs("creator")%>
      </tr>	   
      <tr>
        <td align="right" height="30">修改时间：</td>
        <td class="category"><%=rs("changedate")%>
      </tr>	  
      <tr>
        <td align="right" height="30">修改人：</td>
        <td class="category"><%=rs("changer")%>
      </tr>	            
      <tr>
            <td height="30">&nbsp;</td>
              <td class="category">
            <input type="button" value=" 返回 " onClick="window.history.go(-1);" class="button">&nbsp;&nbsp;&nbsp;&nbsp;        
            </td>
      </tr>   
  </form> 
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

</body>
</html>