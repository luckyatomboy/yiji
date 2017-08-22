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
<title><%=dianming%> - 查看入库单</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script type="text/javascript" src="../js/jquery-3.2.1.js"></script>
<script type="text/javascript" src="../js/jquery-ui.js"></script>
<link href="../style/jquery-ui.css" rel="stylesheet" type="text/css">
<link href="../style/style.css" rel="stylesheet" type="text/css">
<style>
body {
	background-color:#FFFFFF;
}
</style>

<script language=javascript>
function preview() { 
bdhtml=window.document.body.innerHTML; 
sprnstr="<!--startprint-->"; 
eprnstr="<!--endprint-->"; 
prnhtml=bdhtml.substr(bdhtml.indexOf(sprnstr)+17); 
prnhtml=prnhtml.substring(0,prnhtml.indexOf(eprnstr)); 
window.document.body.innerHTML=prnhtml; 
window.print(); 
window.document.body.innerHTML=bdhtml; 
         }
</script>
</HEAD>

<BODY>
<%
if fla32="0" then
%>
<br><center><img src="../images/note.gif" align="absmiddle">&nbsp;<font color="#FF0000">你不具备此权限，请与管理员联系！</font></center>
<%  
  response.end
end if
%>


<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../images/r_1.gif" alt="" /></td>
<td width="100%" background="../images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td>&nbsp;查看入库单资料  ---  <%=rs("stocknumber")%></td>
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
	<form name="form1">
      <tr>
        <td width="20%" align="right" height="30">状态：</td>
        <td width="80%" class="category">
					<%
					sql="select * from stockstatus where category='A'"
					set rs_status=conn.execute(sql)
					%>
					<select name="status" readonly>
						<%do while rs_status.eof=false%>
						<option value="<%=rs_status("status")%>" <%if rs_status("status")=rs("stockstatus") then%>selected="selected"<%end if%>><%=rs_status("description")%></option>
						<%
							rs_status.movenext
							loop
						%>
					</select>				
			  </td>
			</tr>			
      <tr>	  
	    	<td align="right" height="30">参考船期表：</td>
        <td class="category">
					<input name="refshipment" style="width:100px;border:none" readonly value="<%=rs("refshipment")%>">
					&nbsp;&nbsp;&nbsp;项目号
					<input name="refitem" style="width:100px;border:none" readonly value="<%=rs("refitem")%>">
				</td>
      </tr>
      <tr>	  
	    	<td align="right" height="30">入库日期：</td>
        <td class="category">
				<input name="stockdate" id="stockdate" style="width:80px;border:none" readonly value="<%=rs("stockdate")%>">
				</td>
      </tr>     
      <tr>	  
	    	<td align="right" height="30">厂号：</td>
        <td class="category">
					<input name="plant" style="width:100px;border:none" readonly value="<%=rs("plant")%>">
					&nbsp;&nbsp;&nbsp;国别
					<%
					sql="select * from country"
					set rs_country=conn.execute(sql)
					%>
					<select name="country" readonly>
						<%
							do while rs_country.eof=false
						%>
							<option value="<%=rs_country("country")%>" <%if rs_country("country")=rs("country") then%>selected="selected"<%end if%>><%=rs_country("country")%></option>
						<%
							rs_country.movenext
							loop
						%>
					</select>
				&nbsp;&nbsp;&nbsp;合同号	
				<input name="contract" value="<%=rs("contract")%>" style="width:150px;border:none" readonly maxlength="20">
				</td>
      </tr>       
      <tr>	  
	    	<td align="right" height="30">品名：</td>
			<td class="category">
					<%
					sql="select * from material"
					set rs_mat=conn.execute(sql)
					%>
					<select name="material" readonly>
						<%
							do while rs_mat.eof=false
						%>
							<option value="<%=rs_mat("materialname")%>" <%if rs_mat("materialname")=rs("material") then%>selected="selected"<%end if%>><%=rs_mat("materialname")%></option>
						<%
							rs_mat.movenext
							loop
						%>
					</select>	           	
		  		&nbsp;&nbsp;&nbsp;规格
		  		<input name="spec" style="width:100px;border:none" readonly value="<%=rs("spec")%>" readonly>
		  		&nbsp;&nbsp;&nbsp;单价
		  		<input name="price" style="width:100px;border:none" readonly value="<%=rs("price")%>" readonly>&nbsp;/公斤
		  		&nbsp;&nbsp;&nbsp;币种
				<%
				sql="select * from trancurrency"
				set rs_currency=conn.execute(sql)
				%>
				<select name="currency" readonly>
					<%
						do while rs_currency.eof=false
					%>
						<option value="<%=rs_currency("currency")%>" <%if rs_currency("currency")=rs("trancurrency") then%>selected="selected"<%end if%>><%=rs_currency("currency")%></option>
					<%
						rs_currency.movenext
						loop
					%>
				</select>				
			</td>
      </tr>
      <tr>	  
	    	<td align="right" height="30">箱号：</td>
        <td class="category">
		  		<input name="case" style="width:120px;border:none" maxlength="15" value="<%=rs("case")%>" readonly>
		  		&nbsp;&nbsp;&nbsp;箱数
		  		<input name="quantity" style="width:80px;border:none" maxlength="5" value="<%=rs("quantity")%>" readonly >
		  		&nbsp;&nbsp;&nbsp;重量
		  		<input name="weight" style="width:150px;border:none" readonly maxlength="20" value="<%=rs("weight")%>">
				</td>
      </tr>	  
      <tr>	  
	    	<td align="right" height="30">冷库名称：</td>
        <td class="category">
		  		<input name="coldstorage" style="width:100px;border:none" readonly value="<%=rs("coldstorage")%>">
				&nbsp;&nbsp;&nbsp;货主
		  		<input name="customer" style="width:100px;border:none" readonly value="<%=rs("customer")%>">
				</td>
      </tr>
      <tr>	  
	    	<td align="right" height="30">入库数量：</td>
        <td class="category">
		  		<input name="stockqty" style="width:100px;border:none" readonly value="<%=rs("stockqty")%>">
				&nbsp;&nbsp;&nbsp;入库原因
		  		<input name="reason" style="width:100px;border:none" readonly value="<%=rs("reason")%>">				
				&nbsp;&nbsp;&nbsp;剩余数量
		  		<input name="remainqty" style="color:blue;text-decoration:underline;cursor:hand;width:100px" readonly value="<%=rs("remainqty")%>" onClick="JavaScript:window.open('query_remain_stock.asp?refshipment='+refshipment.value+'&refitem='+refitem.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=1200,height=470,top=100,left=20');">				
		</td>
      </tr>
      <tr>	  
	    	<td align="right" height="30">生产日期：</td>
        <td class="category">
		  		<input name="productdate" readonly style="width:80px;border:none" value="<%=rs("productdate")%>">
		  		&nbsp;&nbsp;&nbsp;&nbsp;保质期
				<input name="warranty" style="width:100px;border:none" readonly value="<%=rs("warrantyperiod")%>">
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