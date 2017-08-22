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
<title><%=dianming%> - 查看许可证</title>
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
sql="select * from license where company='"&request("company")&"' and licensetype='"&request("licensetype")&"' and license='"&request("license")&"'"
set rs=conn.execute(sql)
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../images/r_1.gif" alt="" /></td>
<td width="100%" background="../images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td>&nbsp;查看许可证资料</td>
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
        <td width="25%" height="30" align="right">公司名称：</td>
        <td width="75%" class="category"><%=rs("company")%>
        </td>
      </tr>
      <tr>
        <td align="right">类别：</td>
        <td class="category"><%=rs("licensetype")%> 
        </td>
      </tr>      
      <tr>
        <td align="right" height="30">号码：</td>
        <td class="category"><%=rs("license")%>
        </td>
      </tr> 	  
      <tr>
        <td align="right" height="30">年份：</td>
        <td class="category"><input type="text" name="year" style="width:50px;border:none" readonly maxlength="4" value="<%=rs("year")%>"></td>
      </tr>	 
      <tr>
        <td align="right" height="30">国别：</td>
        <td class="category">
					<%
					sql="select * from country order by country"
					set rs_country=conn.execute(sql)
					%>
					<select name="country" readonly>
						<%
							do while rs_country.eof=false
						%>
							<option value="<%=rs_country("country")%>"<%if rs_country("country")=rs("country") then%>selected="selected"<%end if%>><%=rs_country("country")%></option>
						<%
							rs_country.movenext
							loop
						%>
					</select>
        </td>
      </tr>	       
	    <tr>
        <td align="right" height="30">品名：</td>
        <td class="category">
					<%
					sql="select * from material order by materialname"
					set rs_mat=conn.execute(sql)
					%>
					<select name="material" readonly>
						<%
							do while rs_mat.eof=false
						%>
							<option value="<%=rs_mat("materialname")%>"<%if rs_mat("materialname")=rs("material") then%>selected="selected"<%end if%>><%=rs_mat("materialname")%></option>
						<%
							rs_mat.movenext
							loop
						%>
					</select>
        </td>
      </tr> 
      <tr>
        <td align="right" height="30">数量：</td>
        <td class="category"><input type="text" name="quantity" style="width:100px;border:none" readonly value="<%=rs("quantity")%>">&nbsp;&nbsp;吨</td>
      </tr> 	   
      <tr>
        <td align="right" height="30">有效期起始日：</td>
        <td class="category">
		  		<input name="validfrom" readonly style="width:80px;border:none" readonly value="<%=rs("validfrom")%>">
				</td>
      </tr> 	
      <tr>
        <td align="right" height="30">有效期终止日：</td>
        <td class="category">
		  		<input name="validto" readonly style="width:80px;border:none" readonly value="<%=rs("validto")%>">
				</td>
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