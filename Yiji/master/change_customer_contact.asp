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
<title><%=dianming%> - �޸Ŀͻ���ϵ��</title>
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
  sql="select * from locktable where tablename='customercontact' and combinedkey='"&request("customername")&request("contact")&"'"
  set rs_lock=conn.execute(sql)
  if rs_lock.eof = false then
    if rs_lock("username")<>session("redboy_username") then
%>
    <script language="javascript">
    alert("�û�<%=rs_lock("username")%>���ڱ༭�ü�¼�����Ժ����ԣ�");
    window.location.href="shipment.asp";
    </script> 
<%end if
else
    sql="insert into locktable(tablename,combinedkey,status,username,locktime) values('customercontact','"&request("customername")&request("contact")&"','E','"&session("redboy_username")&"',#"&now()&"#)"  
    conn.execute(sql)
end if
%>


<%if request("hid1")="" then%>
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
      <td>&nbsp;�޸Ŀͻ���ϵ������</td>
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
  		<input type="hidden" name="keyword" value="<%=request("keyword")%>">
  		<input type="hidden" name="customername" value="<%=request("customername")%>">
  		<input type="hidden" name="contact" value="<%=request("contact")%>">
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
        <td class="category"><input type="text" name="position" style="width:300px" value="<%=rs("position")%>"></td>
      </tr> 	  
      <tr>
        <td align="right" height="30">�̶��绰��</td>
        <td class="category"><input type="text" name="tel" style="width:200px" value="<%=rs("tel")%>"></td>
      </tr>	 
      <tr>
        <td align="right" height="30">�ֻ���</td>
        <td class="category"><input type="text" name="mobile" style="width:200px" value="<%=rs("mobile")%>"></td>
      </tr>	       
      <tr>
        <td align="right" height="30">���棺</td>
        <td class="category"><input type="text" name="fax" style="width:200px" value="<%=rs("fax")%>"></td>
      </tr>	       
	    <tr>
        <td align="right" height="30">Email��</td>
        <td class="category"><input type="text" name="email" style="width:200px" value="<%=rs("email")%>"></td>
      </tr>        
	    <tr>
        <td align="right" height="30">Skype��</td>
        <td class="category"><input type="text" name="skype" style="width:200px" value="<%=rs("skype")%>"></td>
      </tr>   
	    <tr>
        <td align="right" height="30">΢�ţ�</td>
        <td class="category"><input type="text" name="wechat" style="width:200px" value="<%=rs("wechat")%>"></td>
      </tr>      
	    <tr>
        <td align="right" height="30">QQ��</td>
        <td class="category"><input type="text" name="qq" style="width:200px" value="<%=rs("qq")%>"></td>
      </tr>                  
      <tr>
        <td align="right" height="30">��ע��</td>
        <td class="category"><textarea name="beizhu" cols="70" rows="4"><%=rs("memo")%></textarea></td>
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
		  <input type="submit" value=" ȷ���޸� " onClick="return check()" class="button">&nbsp;&nbsp;&nbsp;&nbsp;
		  <input type="hidden" name="hid1" value="ok">
			<input type="button" value=" �����޸ķ��� " onClick="if (confirm('ȷ��Ҫ�����޸���')) {window.open('delete_lock_table.asp?tablename=customercontact&combinedkey=<%=request("customername")%><%=request("contact")%>'); window.location.href='master.asp';}" class="button">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<%
			if fla6="0" and session("redboy_id")<>"1" then
			else
			%>			
			<input type="button" value=" ɾ�� " onClick="if (confirm('ȷ��Ҫɾ���ÿͻ���ϵ����')) {window.open('delete_customer_contact.asp?customername=<%=request("customername")%>&contact=<%=request("contact")%>')}" class="button"></td>
			<%end if%>		
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
'nowname=request("customername")
nowposition=request("position")
nowtel=request("tel")
nowmobile=request("mobile")
nowfax=request("fax")
nowemail=request("email")
nowskype=request("skype")
nowwechat=request("wechat")
nowqq=request("qq")
nowdes=request("beizhu")
nowkeyword=request("keyword")

set rs=server.createobject("ADODB.RecordSet")
sql="select * from customercontact where customername='"&request("customername")&"' and contact='"&request("contact")&"'"
rs.open sql,conn,1,3
rs("position")=nowposition
rs("tel")=nowtel
rs("mobile")=nowmobile
rs("fax")=nowfax
rs("email")=nowemail
rs("skype")=nowskype
rs("wechat")=nowwechat
rs("qq")=nowqq
rs("memo")=nowdes
rs("changedate")=now
rs("changer")=session("redboy_username")
rs.update
rs.close

%>
<script language="javascript">
alert("�ͻ���ϵ�������޸ĳɹ���")
window.location.href="query_customer_contact.asp?form=<%=request("form")%>&keyword=<%=nowkeyword%>"
</script> 
<%
end if
%>
</body>
</html>