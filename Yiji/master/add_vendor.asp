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
<title><%=dianming%> - 添加供应商</title>
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
//检查必填项
if (document.form1.vendorname.value==""||document.form1.plant1.value=="")
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
      <td>&nbsp;添加供应商(带*号的为必填项)</td>
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
        <td width="75%" class="category"><input type="text" name="vendorname" style="width:200px"> 
        	&nbsp;<font color="#ff0000">*</font></td>
      </tr>    
      <tr>
        <td align="right" height="30">国别：</td>
        <td class="category">
        	<%
					sql="select * from country"
					set rs_country=conn.execute(sql)
					%>
					<select name="country">
						<%
							do while rs_country.eof=false
						%>
							<option value="<%=rs_country("country")%>"><%=rs_country("country")%></option>
						<%
							rs_country.movenext
							loop
						%>
					</select>
        </td>
      </tr>	  
      <tr>
        <td align="right" height="30">公司地址：</td>
        <td class="category"><input type="text" name="address" style="width:300px"></td>
      </tr>       
      <tr>
        <td align="right" height="30">联系电话：</td>
        <td class="category"><input type="text" name="tel" style="width:200px"></td>
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
        <td align="right" height="30">付款方式1：</td>
        <td class="category"><input type="text" name="term1" style="width:100px"></td>
      </tr> 
      <tr>
        <td align="right" height="30">付款方式2：</td>
        <td class="category"><input type="text" name="term2" style="width:100px"></td>
      </tr>       
<!--added on 2017/2/4  begin-->
      <tr>
        <td align="right" height="30">厂号1：</td>
        <td class="category"><input type="text" name="plant1" style="width:60px"> 
        	<font color="#ff0000">*</font>
        	&nbsp;&nbsp;&nbsp;厂号2 <input type="text" name="plant2" style="width:60px"> 
        	&nbsp;&nbsp;&nbsp;厂号3 <input type="text" name="plant3" style="width:60px"> 
        	&nbsp;&nbsp;&nbsp;厂号4 <input type="text" name="plant4" style="width:60px"> 
        </td>
      </tr> 	      
<!--end -->
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
nowvendorname=request("vendorname")
nowcountry=request("country")
nowtel=request("tel")
nowaddress=request("address")
nowfax=request("fax")
nowemail=request("email")
'nowpartner=request("partner")
nowplant1=request("plant1")
nowplant2=request("plant2")
nowplant3=request("plant3")
nowplant4=request("plant4")
nowterm1=request("term1")
nowterm2=request("term2")
nowbeizhu=request("beizhu")
sql="select * from Vendor where Vendorname='"&nowvendorname&"'"
set rs=conn.execute(sql)
if rs.eof=false then
%>
<script language="javascript">
alert("您输入的供应商已经存在，请重新输入！")
window.history.go(-1)
</script> 
<%
  response.end
end if
rs.close

'检查工厂是否重复'
for i = 1 to 4
  if request("plant"&i) <> "" then
    sql="select * from plant where plantid = '"&request("plant"&i)&"' and country = '"&nowcountry&"'"
    set rs=conn.execute(sql)
    if rs.eof=false then
    %>
    <script language="javascript">
    alert("<%=nowcountry%>工厂<%=rs("plantid")%>已被供应商<%=rs("vendor")%>占用！");
    window.history.go(-1)
    </script> 
    <%
      response.end
    end if
    rs.close
  end if
next 

'添加供应商数据'
sql="insert into Vendor(VendorName,Country,Address,Tel,Fax,Email,Plant1,Plant2,Plant3,Plant4,Memo,Term1,Term2,CreateDate,Creator) values('"&nowvendorname&"','"&nowcountry&"','"&nowaddress&"','"&nowtel&"','"&nowfax&"','"&nowemail&"','"&nowplant1&"','"&nowplant2&"','"&nowplant3&"','"&nowplant4&"','"&nowbeizhu&"','"&nowterm1&"','"&nowterm2&"',#"&now()&"#,'"&session("redboy_username")&"')"
conn.execute(sql)

'添加工厂数据
for i = 1 to 4
  if request("plant"&i) <> "" then  
    sql="insert into plant(plantid,Country,vendor,CreateDate,Creator) values('"&request("plant"&i)&"','"&nowcountry&"','"&nowvendorname&"',#"&now()&"#,'"&session("redboy_username")&"')"
    conn.execute(sql)
  end if
next
%>
<script language="javascript">
alert("供应商添加成功！")
window.location.href="master.asp"
</script> 
<%
end if
%>
</body>
</html>