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
<title><%=dianming%> - 添加供应商联系人</title>
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
<%if request("hid1")="" then%> 
<script language="javascript">
function check()
{
if (document.form1.vendor.value=="" || document.form1.contact.value=="" || document.form1.userid.value=="")
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
      <td>&nbsp;添加供应商联系人(带*号的为必填项)</td>
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
        <td width="25%" height="30" align="right">供应商名称：</td>
        <td width="75%" class="category">
        	<%
					sql="select * from vendor"
					set rs_vendor=conn.execute(sql)
					%>
					<select name="vendor">
						<%
							do while rs_vendor.eof=false
						%>
							<option value="<%=rs_vendor("vendorname")%>"><%=rs_vendor("vendorname")%></option>
						<%
							rs_vendor.movenext
							loop
						%>
					</select>        	
        	&nbsp;<font color="#ff0000">*</font></td>
      </tr>    
      <tr>
        <td align="right" height="30">联系人：</td>
        <td class="category"><input type="text" name="contact" style="width:100px">
        	&nbsp;<font color="#ff0000">*</font>
        </td>
      </tr> 
      <tr>
        <td align="right" height="30">职位：</td>
        <td class="category"><input type="text" name="position" style="width:100px"></td>
      </tr>           
      <tr>
        <td align="right" height="30">固定电话：</td>
        <td class="category"><input type="text" name="tel" style="width:150px"></td>
      </tr>	 
      <tr>
        <td align="right" height="30">手机：</td>
        <td class="category"><input type="text" name="mobile" style="width:150px"></td>
      </tr>	   
      <tr>
        <td align="right" height="30">传真：</td>
        <td class="category"><input type="text" name="fax" style="width:200px"></td>
      </tr>	       
	    <tr>
        <td align="right" height="30">Email：</td>
        <td class="category"><input type="text" name="email" style="width:200px"></td>
      </tr> 
      <tr>
        <td align="right" height="30">Skype：</td>
        <td class="category"><input type="text" name="Skype" style="width:100px"></td>
      </tr>     
      <tr>
        <td align="right" height="30">微信：</td>
        <td class="category"><input type="text" name="Wechat" style="width:100px"></td>
      </tr>   
      <tr>
        <td align="right" height="30">QQ：</td>
        <td class="category"><input type="text" name="qq" style="width:100px"></td>
      </tr>   
      <tr>
        <td align="right" height="30">备注：</td>
        <td class="category"><textarea name="beizhu" cols="70" rows="4"></textarea></td>
      </tr>	  
      <tr>
	    <td height="30">&nbsp;</td>
        <td class="category">
		  <input type="submit" value=" 确认添加 " onClick="return check()" class="button">&nbsp;&nbsp;&nbsp;&nbsp;
		  <input type="hidden" name="hid1" value="ok">
		  <input type="reset" value=" 重新填写 " class="button">		</td>
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
nowvendor=request("vendor")
nowcontact=request("contact")
nowposition=request("position")
nowtel=request("tel")
nowmobile=request("mobile")
nowfax=request("fax")
nowemail=request("email")
nowskype=request("skype")
nowwechat=request("wechat")
nowqq=request("qq")
nowbeizhu=request("beizhu")
sql="select * from VendorContact where Vendorname='"&nowvendor&"' and contact='"&nowcontact&"'"
set rs=conn.execute(sql)
if rs.eof=false then
%>
<script language="javascript">
alert("您输入的供应商联系人已经存在，请重新输入！")
window.history.go(-1)
</script> 
<%
  response.end
end if

sql="insert into VendorContact(VendorName,Contact,Position,Tel,Mobile,Fax,Email,Skype,Wechat,QQ,Memo,CreateDate,Creator) values('"&nowvendor&"','"&nowcontact&"','"&nowposition&"','"&nowtel&"','"&nowmobile&"','"&nowfax&"','"&nowemail&"','"&nowskype&"','"&nowwechat&"','"&nowqq&"','"&nowbeizhu&"',#"&now()&"#,'"&session("redboy_username")&"')"
conn.execute(sql)
%>
<script language="javascript">
alert("供应商联系人添加成功！")
window.location.href="master.asp"
</script> 
<%
end if
%>
</body>
</html>