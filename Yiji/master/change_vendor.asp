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
<title><%=dianming%> - 修改供应商</title>
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
<br><center><img src="../images/note.gif" align="absmiddle">&nbsp;<font color="#FF0000">你不具备此权限，请与管理员联系！</font></center>
<%  
  response.end
end if
%>

<%
  sql="select * from locktable where tablename='vendor' and combinedkey='"&request("vendorname")&"'"
  set rs_lock=conn.execute(sql)
  if rs_lock.eof = false then
%>
    <script language="javascript">
    alert("用户<%=rs_lock("username")%>正在编辑该记录！请稍后再试！");
    window.location.href="master.asp";
    </script> 
<%else
    sql="insert into locktable(tablename,combinedkey,status,username,locktime) values('vendor','"&request("vendorname")&"','E','"&session("redboy_username")&"',#"&now()&"#)"  
    conn.execute(sql)
end if
%>

<%if request("hid1")="" then

sql="select * from vendor where vendorname='"&request("vendorname")&"'"
set rs=conn.execute(sql)
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../images/r_1.gif" alt="" /></td>
<td width="100%" background="../images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td>&nbsp;修改供应商资料</td>
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
  		<input type="hidden" name="vendorname" value="<%=request("vendorname")%>">
      <tr>
        <td width="25%" height="30" align="right">供应商名称：</td>
        <td width="75%" class="category"><%=rs("vendorname")%>
        	</td>
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
        <td align="right" height="30">付款方式1：</td>
        <td class="category"><input type="text" name="term1" style="width:100px" value="<%=rs("term1")%>"></td>
      </tr> 	
      <tr>
        <td align="right" height="30">付款方式2：</td>
        <td class="category"><input type="text" name="term2" style="width:100px" value="<%=rs("term2")%>"></td>
      </tr> 	    
<!--added on 2017/2/4  begin-->
      <tr>
        <td align="right" height="30">厂号1：</td>
        <td class="category"><input type="text" name="plant1" style="width:60px" value="<%=rs("plant1")%>"> 
        	&nbsp;&nbsp;&nbsp;厂号2 <input type="text" name="plant2" style="width:60px" value="<%=rs("plant2")%>"> 
        	&nbsp;&nbsp;&nbsp;厂号3 <input type="text" name="plant3" style="width:60px" value="<%=rs("plant3")%>"> 
        	&nbsp;&nbsp;&nbsp;厂号4 <input type="text" name="plant4" style="width:60px" value="<%=rs("plant4")%>"> 
        </td>
      </tr> 	      
<!--end -->
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
		  <input type="submit" value=" 确认修改 " class="button">&nbsp;&nbsp;&nbsp;&nbsp;
		  <input type="hidden" name="hid1" value="ok">
      <input type="hidden" name="hid2" value="vendor">
			<input type="button" value=" 放弃修改返回 " onClick="window.open('delete_lock_table.asp?tablename='<%=request("hid2")%>'&combinedkey='<%=request("vendorname")%>'')" class="button">&nbsp;&nbsp;&nbsp;&nbsp;
			<%
			if fla35="0" and session("redboy_id")<>"1" then
			else
			%>			
			<input type="button" value=" 删除 " onClick="if (confirm('确定要删除该供应商吗？')) {window.open('delete_vendor.asp?vendorname=<%=request("vendorname")%>');}" class="button"></td>
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
nowname=request("vendorname")
nowaddress=request("address")
nowtel=request("tel")
nowfax=request("fax")
nowemail=request("email")
'nowpartner=request("partner")
nowplant1=request("plant1")
nowplant2=request("plant2")
nowplant3=request("plant3")
nowplant4=request("plant4")
nowdes=request("beizhu")
nowkeyword=request("keyword")

set rs=server.createobject("ADODB.RecordSet")
sql="select * from vendor where vendorname='"&request("vendorname")&"'"
rs.open sql,conn,1,3
rs("address")=nowaddress
rs("tel")=nowtel
rs("fax")=nowfax
rs("email")=nowemail
rs("plant1")=nowplant1
rs("plant2")=nowplant2
rs("plant3")=nowplant3
rs("plant4")=nowplant4
rs("memo")=nowdes
rs("changedate")=now
rs("changer")=session("redboy_username")
rs.update
rs.close

sql="delete from locktable where tablename='vendor' and combinedkey="&request("vendorname")
conn.execute(sql)

%>
<script language="javascript">
alert("客户资料修改成功！")
window.location.href="query_vendor.asp?form=<%=request("form")%>&keyword=<%=nowkeyword%>"
</script> 
<%
end if
%>
</body>
</html>