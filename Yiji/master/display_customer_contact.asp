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
<title><%=dianming%> - �鿴�ͻ���ϵ��</title>
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
<br><center><img src="../images/note.gif" align="absmiddle">&nbsp;<font color="#FF0000">�㲻�߱���Ȩ�ޣ��������Ա��ϵ��</font></center>
<%  
  response.end
end if
%>

<%
sql="select * from customercontact where customername='"&request("customername")&"' and contact='"&request("contact")&"'"
set rs=conn.execute(sql)
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../images/r_1.gif" alt="" /></td>
<td width="100%" background="../images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td>&nbsp;�鿴�ͻ���ϵ������</td>
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
        <td width="25%" height="30" align="right">�ͻ����ƣ�</td>
        <td width="75%" class="category"><%=rs("customername")%>
        	</td>
      </tr>    
      <tr>
        <td width="25%" height="30" align="right">��ϵ�ˣ�</td>
        <td width="75%" class="category"><%=rs("contact")%>
        	</td>
      </tr>        
      <tr>
        <td align="right" height="30">ְλ��</td>
        <td class="category"><input type="text" name="position" style="width:300px;border:none" readonly value="<%=rs("position")%>"></td>
      </tr> 	  
      <tr>
        <td align="right" height="30">�̶��绰��</td>
        <td class="category"><input type="text" name="tel" style="width:200px;border:none" readonly value="<%=rs("tel")%>"></td>
      </tr>	 
      <tr>
        <td align="right" height="30">�ֻ���</td>
        <td class="category"><input type="text" name="mobile" style="width:200px;border:none" readonly value="<%=rs("mobile")%>"></td>
      </tr>	       
      <tr>
        <td align="right" height="30">���棺</td>
        <td class="category"><input type="text" name="fax" style="width:200px;border:none" readonly value="<%=rs("fax")%>"></td>
      </tr>	       
	    <tr>
        <td align="right" height="30">Email��</td>
        <td class="category"><input type="text" name="email" style="width:200px;border:none" readonly value="<%=rs("email")%>"></td>
      </tr>        
	    <tr>
        <td align="right" height="30">Skype��</td>
        <td class="category"><input type="text" name="skype" style="width:200px;border:none" readonly value="<%=rs("skype")%>"></td>
      </tr>   
	    <tr>
        <td align="right" height="30">΢�ţ�</td>
        <td class="category"><input type="text" name="wechat" style="width:200px;border:none" readonly value="<%=rs("wechat")%>"></td>
      </tr>      
	    <tr>
        <td align="right" height="30">QQ��</td>
        <td class="category"><input type="text" name="qq" style="width:200px;border:none" readonly value="<%=rs("qq")%>"></td>
      </tr>                  
      <tr>
        <td align="right" height="30">��ע��</td>
        <td class="category"><textarea name="beizhu" cols="70" rows="4" readonly><%=rs("memo")%></textarea></td>
      </tr>	   
      <tr>
        <td align="right" height="30">����ʱ�䣺</td>
        <td class="category"><%=rs("createdate")%>
      </tr>	  
      <tr>
        <td align="right" height="30">�����ˣ�</td>
        <td class="category"><%=rs("creator")%>
      </tr>	   
      <tr>
        <td align="right" height="30">�޸�ʱ�䣺</td>
        <td class="category"><%=rs("changedate")%>
      </tr>	  
      <tr>
        <td align="right" height="30">�޸��ˣ�</td>
        <td class="category"><%=rs("changer")%>
      </tr>	            
      <tr>
            <td height="30">&nbsp;</td>
              <td class="category">
            <input type="button" value=" ���� " onClick="window.history.go(-1);" class="button">&nbsp;&nbsp;&nbsp;&nbsp;        
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