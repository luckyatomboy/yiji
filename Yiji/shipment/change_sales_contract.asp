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
<title><%=dianming%> - 修改订货合同</title>
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
<!--if fla1="0" and session("redboy_id")<>"1" then-->
if fla27="0" then
%>
<br><center><img src="../images/note.gif" align="absmiddle">&nbsp;<font color="#FF0000">你不具备此权限，请与管理员联系！</font></center>
<%  
  response.end
end if
%>

<%
  sql="select * from locktable where tablename='SalesContract' and combinedkey='"&request("ContractNum")&"'"
  set rs_lock=conn.execute(sql)
  if rs_lock.eof = false then
%>
    <script language="javascript">
    alert("用户<%=rs_lock("username")%>正在编辑该记录！请稍后再试！");
    window.location.href="shipment.asp";
    </script> 
<%else
    sql="insert into locktable(tablename,combinedkey,status,username,locktime) values('salescontract','"&request("ContractNum")&"','E','"&session("redboy_username")&"',#"&now()&"#)"  
    conn.execute(sql)
end if
%>

<%
if request("hid1")="ok" then
nowcategory=request("category")
nowowncompany=request("owncompany")
nowcustomer=request("customer")
nowstatus=request("status")
nowrefshipment=request("refshipment")
nowrefitem=request("refitem")
nowcountry=request("guobie")
if request("material")<>"" then
	nowmaterial=request("material")
	nowspec=request("spec")
	nowprice=request("price")
else
	nowmaterial=""
	nowspec=""
	nowprice=0	
end if
nowplant=request("plant")
'nowcasenumber=request("casenumber")
'nowcase=request("case")
if request("weight")<>"" then
	nowweight=request("weight")
else
	nowweight=0
end if
if request("quantity")<>"" then
	nowquantity=request("quantity")
else
	nowquantity=0
end if

if request("boarddate")<>"" then
	nowboarddate=request("boarddate")
else
	nowboarddate=date()
end if
nowpackage=request("package")
nowstorage=request("coldstorage")
nowdeliveryloc=request("deliveryloc")
nowdeliveryport=request("deliveryport")

sql="select * from SalesContract where ContractNum="&request("ContractNum")
set rs=server.createobject("ADODB.RecordSet")
rs.open sql,conn,1,3

rs("category")=nowcategory
rs("owncompany")=nowowncompany
rs("customer")=nowcustomer
rs("status")=nowstatus
rs("quantity")=nowquantity
rs("weight")=nowweight
rs("price")=nowprice
rs.update
rs.close

'如果有入库单，更新剩余库存数量'
sql="select * from stockdocument where refshipment="&nowrefshipment&" and refitem="&nowrefitem
set rs=server.createobject("ADODB.RecordSet")
rs.open sql,conn,1,3

nowremainqty=rs("remainqty")-nowquantity
rs("remainqty")=nowremainqty
rs("changedate")=now()
rs("changer")=session("redboy_username")
rs.update
rs.close

sql="delete from locktable where tablename='SalesContract' and combinedkey="&request("ContractNum")
conn.execute(sql)

%>
<script language="javascript">
//alert(&nowshipment&)
alert("订货合同修改成功！")
window.location.href="shipment.asp"
</script>

<%
else
%>

<%
sql="select * from SalesContract where ContractNum="&request("ContractNum")
set rs=conn.execute(sql)
%>

<script language="javascript">
function check1(){
<!-- 检查品名 -->
if (document.form1.material.value=="")
	{
	alert("请输入品名！");
	return false;
	}
<!-- 检查单价 -->
if (document.form1.price.value=="0" || document.form1.price.value=="")
	{
	alert("请输入单价！");
	return false;
	}	
<!-- 检查日期 -->
if (document.form1.category.value=="B" && document.form1.boarddate.value=="")
	{
	alert("期货合同请输入预计到港期！");
	return false;
	}
}
</script>

<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../images/r_1.gif" alt="" /></td>
<td width="100%" background="../images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td>&nbsp;订货合同&nbsp;<%=rs("ContractNum")%> </td>
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
	  <input type="hidden" name="ContractNum" value="<%=request("ContractNum")%>">
      <tr>
        <td align="right" height="30">合同类型：</td>
        <td class="category">
			<select name="category">
				<option value="A" <%if rs("status")="A" then%>selected="selected"<%end if%>>现货</option>
				<option value="B" <%if rs("status")="B" then%>selected="selected"<%end if%>>期货</option>
			</select>				
	    </td>
	  </tr>	  
      <tr>
        <td align="right" height="30">我方公司：</td>
        <td class="category">
			<%
			sql="select * from owncompany order by company"
			set rs_owncompany=conn.execute(sql)
			%>		  		
			<select name="owncompany">			
				<%
					do while rs_owncompany.eof=false
				%>
					<option value="<%=rs_owncompany("company")%>" <%if rs("owncompany")=rs_owncompany("company") then%>selected="selected"<%end if%>><%=rs_owncompany("company")%></option>
				<%
					rs_owncompany.movenext
					loop
					rs_owncompany.close
				%>								
			</select>			
			&nbsp;&nbsp;&nbsp;&nbsp;客户
			<%
			sql="select * from customer order by customername"
			set rs_customer=conn.execute(sql)
			%>		  		
			<select name="customer">			
				<%
					do while rs_customer.eof=false
				%>
					<option value="<%=rs_customer("customername")%>" <%if rs("customer")=rs_customer("customername") then%>selected="selected"<%end if%>><%=rs_customer("customername")%></option>
				<%
					rs_customer.movenext
					loop
					rs_customer.close
				%>								
			</select>	
	    </td>    		
	  </tr>	  
      <tr>
        <td width="20%" align="right" height="30">状态：</td>
        <td width="80%" class="category">
			<%
			sql="select * from contractstatus"
			set rs_status=conn.execute(sql)
			%>
			<select name="status">
				<%do while rs_status.eof=false%>
				<option value="<%=rs_status("status")%>" <%if rs("status")=rs_status("status") then%>selected="selected"<%end if%>><%=rs_status("description")%></option>
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
					<input name="refshipment" readonly style="cursor:hand;width:100px" value="<%=rs("refshipment")%>"> 
					&nbsp;&nbsp;&nbsp;项目号	
					<input name="refitem" readonly style="width:50px" value="<%=rs("refitem")%>"> 
				</td>
      </tr>   
      <tr>	  
	    <td align="right" height="30">厂号：</td>
        <td class="category">
			<input name="plant" style="width:100px" readonly value="<%=rs("plant")%>">
			&nbsp;&nbsp;&nbsp;国别
			<input name="guobie" style="width:100px" readonly value="<%=rs("country")%>">
<!--			&nbsp;&nbsp;&nbsp;我司编号	
			<input name="shipmentnum" style="width:150px" maxlength="20" readonly> -->
		</td>
      </tr>       
      <tr>	  
	    <td align="right" height="30">品名：</td>
		<td class="category">
          		<input name="material" style="width:100px" readonly value="<%=rs("material")%>">
		  		&nbsp;&nbsp;&nbsp;规格
		  		<input name="spec" style="width:100px" readonly value="<%=rs("spec")%>">
		  		&nbsp;&nbsp;&nbsp;包装
		  		<input name="package" style="width:200px" value="<%=rs("package")%>">		  		
		</td>
      </tr>
      <tr>	  
	    <td align="right" height="30">件数：</td>
        <td class="category">
		  		<input name="quantity" style="width:60px" maxlength="5" value="<%=rs("quantity")%>" onKeyPress="javascript:CheckNum();"  onKeyUp="this.value=this.value.replace(/[^\d.]/g,'')">
		  		&nbsp;&nbsp;&nbsp;重量
		  		<input name="weight" style="width:50px" maxlength="20" value="<%=rs("weight")%>">吨
		  		&nbsp;&nbsp;&nbsp;单价
		  		<input name="price" style="width:50px" value="<%=rs("price")%>" onKeyPress="javascript:CheckNum();" onKeyUp="this.value=this.value.replace(/[^\d.]/g,'')">元/吨		  		
		</td>
      </tr>	  
      <tr>	  
	    <td align="right" height="30">存放冷库：</td>
        <td class="category">
		  		<input name="coldstorage" style="width:100px" readonly value="<%=rs("coldstorage")%>">
				&nbsp;&nbsp;&nbsp;交货地
		  		<input name="deliveryloc" style="width:100px" value="<%=rs("deliveryloc")%>">
				</td>
      </tr>
      <tr>	  
	    <td align="right" height="30">预计到港期：</td>
        <td class="category">
		  		<input name="boarddate" style="width:80px" value="<%=rs("boarddate")%>">
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=boarddate&oldDate='+boarddate.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">
				&nbsp;&nbsp;&nbsp;交货港
		  		<input name="deliveryport" style="width:100px" value="<%=rs("deliveryport")%>">							
				</td>
      </tr>
      <tr>
	    <td height="30">&nbsp;</td>
        <td class="category">
		  <input type="submit" value=" 确认修改 " onClick="return check1()" class="button">
		  <input type="hidden" name="hid1" value="ok">
		  <input type="button" value=" 放弃修改返回 " onClick="if (confirm('确定要放弃修改吗？')) {window.open('../master/delete_lock_table.asp?tablename=salescontract&combinedkey=<%=request("contractnum")%>');window.history.go(-2);}" class="button">
			<%
			if fla7="0" and session("redboy_id")<>"1" then
			else
			%>			
			<input type="button" value=" 删除 " onClick="if (confirm('确定要删除该订货合同吗？')) {window.open('delete_sales_contract.asp?status=<%=rs("status")%>&ContractNum=<%=request("ContractNum")%>'); window.history.go(-2);}" class="button"></td>
			<%end if%>			  
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
<%end if%>
</body>
</html>