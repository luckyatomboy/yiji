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
<title><%=dianming%> - 创建订货合同</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script type="text/javascript" src="../js/jquery-3.2.1.min.js"></script>
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
if request("hid1")="ok" then%>



<%
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

if nowcategory="A" then
	sql="select top 1 * from SalesContract where ContractNum like '1%' order by ContractNum desc"
else
	sql="select top 1 * from SalesContract where ContractNum like '2%' order by ContractNum desc"
end if

set rs_count=conn.execute(sql)

if rs_count.bof and rs_count.eof then
	if nowcategory="A" then
		nowcontract="10000001"
	else
		nowcontract="20000001"
	end if
else
	rs_count.movefirst
	nowcontract=rs_count("ContractNum") + 1
end if	

'创建订货合同'
sql="insert into SalesContract(ContractNum,category,status,owncompany,customer,country,plant,material,spec,package,quantity,weight,price,coldstorage,deliveryloc,boarddate,deliveryport,refshipment,refitem,createdate,creator)"
sql=sql&" values("&nowcontract&",'"&nowcategory&"','"&nowstatus&"','"&nowowncompany&"','"&nowcustomer&"','"&nowcountry&"','"&nowplant&"','"&nowmaterial&"','"&nowspec&"','"&nowpackage&"',"&nowquantity&","&nowweight&","&nowprice&",'"&nowstorage&"','"&nowdeliveryloc&"',#"&nowboarddate&"#,'"&nowdeliveryport&"','"&nowrefshipment&"','"&nowrefitem&"',#"&now()&"#,'"&session("redboy_username")&"')"
conn.execute(sql)

'如果有入库单，更新剩余库存数量'
sql="select * from stockdocument where refshipment="&nowrefshipment&" and refitem="&nowrefitem
set rs=server.createobject("ADODB.RecordSet")
rs.open sql,conn,1,3

if rs.eof=false then
nowremainqty=rs("remainqty")-nowquantity
rs("remainqty")=nowremainqty
rs("changedate")=now()
rs("changer")=session("redboy_username")
rs.update
rs.close
end if

%>
<script language="javascript">
//alert(&nowshipment&)
alert("订货合同录入成功！")
window.location.href="shipment.asp"
</script>

<%
else
%>

<script LANGUAGE="JAVASCRIPT">

function check_sales_contract(){
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
<!-- 检查数量 -->
if (document.form1.quantity.value=="0" || document.form1.quantity.value=="")
	{
	alert("请输入数量！");
	return false;
	}		
<!-- 检查日期 -->
if (document.form1.category.value=="B" && document.form1.boarddate.value=="")
	{
	alert("期货合同请输入预计到港期！");
	return false;
	}

<!--如果是现货合同，检查剩余库存-->
if (document.form1.category.value=="A"){
	if (window.XMLHttpRequest)
	  {// 针对 IE7+, Firefox, Chrome, Opera, Safari 的代码
	 	 xmlhttp=new XMLHttpRequest();
	  }
	else
	  {// 针对 IE6, IE5 的代码
	  	xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
	  }
	xmlhttp.onreadystatechange=function()
	  {
	  if (xmlhttp.readyState==4 && xmlhttp.status==200)
	    {
	    	remainqty=xmlhttp.responseText;
	    //document.getElementByName("remainqty").value=xmlhttp.responseText;
	    	if (remainqty=="e"){
	    		alert("没有入库单，不能创建现货合同");
	    		result="false";
	    	}else if (parseInt(remainqty)<parseInt(document.form1.quantity.value)){
	    		alert("剩余库存为"+remainqty+"，请重新输入数量");
	    		result="false";
	    	}else{
	    		result="true";
	    	}
	    }
	  }
	xmlhttp.open("GET","check_remain_qty.asp?refshipment="+document.form1.refshipment.value+"&refitem="+document.form1.refitem.value,false);
	xmlhttp.send();
//	window.open("check_remain_qty.asp?refshipment="+document.form1.refshipment.value+"&refitem="+document.form1.refitem.value,"","directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590");
	
	if (result=="false"){
		return false;
	}
	}
}

</script>

<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../images/r_1.gif" alt="" /></td>
<td width="100%" background="../images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td>&nbsp;订货合同 </td>
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
	  <input type="hidden" name="remainqty" value="<%=request("remainqty")%>">
      <tr>
        <td align="right" height="30">合同类型：</td>
        <td class="category">
			<select name="category">
				<option value="A">现货</option>
				<option value="B">期货</option>
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
					<option value="<%=rs_owncompany("company")%>"><%=rs_owncompany("company")%></option>
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
					<option value="<%=rs_customer("customername")%>"><%=rs_customer("customername")%></option>
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
				<option value="<%=rs_status("status")%>"><%=rs_status("description")%></option>
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
					<input name="refshipment" readonly style="cursor:hand;width:100px" value="单击选择船期表项目" onClick="JavaScript:window.open('query_shipment_new.asp?queryform=form1&refship=refshipment&refitem=refitem&plant=plant&material=material&guobie=guobie&spec=spec&package=package&quantity=quantity&weight=weight&coldstorage=coldstorage&boarddate=boarddate&deliveryport=deliveryport','','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=1200,height=470,top=100,left=20');"> &nbsp;<font color="#ff0000">*</font>
					&nbsp;&nbsp;&nbsp;项目号	
					<input name="refitem" readonly style="width:50px"
				</td>
      </tr>   
      <tr>	  
	    <td align="right" height="30">厂号：</td>
        <td class="category">
			<input name="plant" style="width:100px" readonly>
			&nbsp;&nbsp;&nbsp;国别
			<input name="guobie" style="width:100px" readonly>
<!--			&nbsp;&nbsp;&nbsp;我司编号	
			<input name="shipmentnum" style="width:150px" maxlength="20" readonly> -->
		</td>
      </tr>       
      <tr>	  
	    <td align="right" height="30">品名：</td>
		<td class="category">
          		<input name="material" style="width:100px" readonly>
		  		&nbsp;&nbsp;&nbsp;规格
		  		<input name="spec" style="width:100px" readonly>
		  		&nbsp;&nbsp;&nbsp;包装
		  		<input name="package" style="width:200px">		  		
		</td>
      </tr>
      <tr>	  
	    <td align="right" height="30">件数：</td>
        <td class="category">
		  		<input name="quantity" style="width:60px" maxlength="5"  <% if nowquantity<>"" then%>value="<%=nowquantity%>"<%end if%> onKeyPress="javascript:CheckNum();"  onKeyUp="this.value=this.value.replace(/[^\d.]/g,'')">
		  		&nbsp;&nbsp;&nbsp;重量
		  		<input name="weight" style="width:50px" maxlength="20" <% if nowweight<>"" then%>value="<%=nowweight%>"<%end if%>>吨
		  		&nbsp;&nbsp;&nbsp;单价
		  		<input name="price" style="width:50px" value=0 onKeyPress="javascript:CheckNum();" onKeyUp="this.value=this.value.replace(/[^\d.]/g,'')">元/吨		  		
		</td>
      </tr>	  
      <tr>	  
	    <td align="right" height="30">存放冷库：</td>
        <td class="category">
		  		<input name="coldstorage" style="width:100px" >
				&nbsp;&nbsp;&nbsp;交货地
		  		<input name="deliveryloc" value="上海" style="width:100px">
				</td>
      </tr>
      <tr>	  
	    <td align="right" height="30">预计到港期：</td>
        <td class="category">
		  		<input name="boarddate" style="width:80px">
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=boarddate&oldDate='+boarddate.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">
				&nbsp;&nbsp;&nbsp;交货港
		  		<input name="deliveryport" style="width:100px">							
				</td>
      </tr>
      <tr>
	    <td height="30">&nbsp;</td>
        <td class="category">
		  <input type="submit" value=" 确认录入 " onClick="return check_sales_contract()" class="button">
		  <input type="hidden" name="hid1" value="ok">
		  <input type="reset" value=" 重新填写 " class="button">
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