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
<title><%=dianming%> - 修改客户</title>
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
  sql="select * from locktable where tablename='customer' and combinedkey='"&request("customername")&"'"
  set rs_lock=conn.execute(sql)
  if rs_lock.eof = false then
    if rs_lock("username")<>session("redboy_username") then
%>
    <script language="javascript">
    alert("用户<%=rs_lock("username")%>正在编辑该记录！请稍后再试！");
    window.location.href="master.asp";
    </script> 
<%end if
else
    sql="insert into locktable(tablename,combinedkey,status,username,locktime) values('customer','"&request("customername")&"','E','"&session("redboy_username")&"',#"&now()&"#)"  
    conn.execute(sql)
end if
%>

<%if request("hid1")="" then%>

<script language="javascript">
function check()
{
if (document.form1.username.value==""||document.form1.userid.value=="")
{
alert("有*号的必须填写！");
return false;
}
}


</script>

<%
sql="select * from customer where customername='"&request("customername")&"'"
set rs=conn.execute(sql)
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../images/r_1.gif" alt="" /></td>
<td width="100%" background="../images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td>&nbsp;修改客户资料</td>
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
      <tr>
        <td width="25%" height="30" align="right">客户简称：</td>
        <td width="75%" class="category"><%=rs("customername")%></td>
      </tr>    
      <tr>
        <td height="30" align="right">客户全称：</td>
        <td class="category"><input type="text" name="fullname" style="width:200px" value="<%=rs("fullname")%>">
        &nbsp;<font color="#ff0000">*</font></td>
      </tr>        
      <tr>
        <td align="right" height="30">公司地址：</td>
        <td class="category"><input type="text" name="address" style="width:300px" value="<%=rs("address")%>"></td>
      </tr> 	  
      <tr>
        <td align="right" height="30">联系电话：</td>
        <td class="category"><input type="text" name="tel" style="width:200px" value="<%=rs("tel")%>"></td>
      </tr>	 
      <tr>
        <td align="right" height="30">传真：</td>
        <td class="category"><input type="text" name="fax" style="width:200px" value="<%=rs("fax")%>"></td>
      </tr>	       
	    <tr>
        <td align="right" height="30">Email：</td>
        <td class="category"><input type="text" name="email" style="width:200px" value="<%=rs("email")%>"></td>
      </tr>        
      <tr>
        <td align="right" height="30">备注：</td>
        <td class="category"><textarea name="beizhu" cols="70" rows="4"><%=rs("memo")%></textarea></td>
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
		  <input type="submit" value=" 确认修改 " onClick="return check()" class="button">&nbsp;&nbsp;&nbsp;&nbsp;
		  <input type="hidden" name="hid1" value="ok">
		  <input type="hidden" name="hid2" value="customer">
			<input type="button" value=" 放弃修改返回 " onClick="if (confirm('确定要放弃修改吗？')) {window.open('delete_lock_table.asp?combinedkey=<%=request("customername")%>'); window.history.go(-2);}" class="button>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<%
			if fla7="0" and session("redboy_id")<>"1" then
			else
			%>			
			<input type="button" value=" 删除 " onClick="if (confirm('确定要删除该客户吗？')) {window.open('delete_customer.asp?customername=<%=request("customername")%>'); window.history.go(-2);}" class="button"></td>
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
%>

<%
nowname=request("customername")
nowfullname=request("fullname")
nowaddress=request("address")
nowtel=request("tel")
nowfax=request("fax")
nowemail=request("email")
nowdes=request("beizhu")
nowkeyword=request("keyword")

set rs=server.createobject("ADODB.RecordSet")
sql="select * from customer where customername='"&request("customername")&"'"
rs.open sql,conn,1,3
rs("fullname")=nowfullname
rs("address")=nowaddress
rs("tel")=nowtel
rs("fax")=nowfax
rs("email")=nowemail
rs("memo")=nowdes
rs("changedate")=now
rs("changer")=session("redboy_username")
rs.update
rs.close

sql="delete from locktable where tablename='customer' and combinedkey='"&request("customername")&"'"
conn.execute(sql)
%>
<script language="javascript">
alert("客户资料修改成功！")
//window.location.href="query_customer.asp?form=<%=request("form")%>&keyword=<%=nowkeyword%>"
window.location.href="master.asp"
</script> 
<%
end if
%>
</body>
</html>