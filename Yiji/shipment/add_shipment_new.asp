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
<title><%=dianming%> - 船期表</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--<meta http-equiv="Content-Type" content="application/x-www-form-urlencoded; charset=gb2312">-->
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
if fla1="0" then
%>
<br><center><img src="../images/note.gif" align="absmiddle">&nbsp;<font color="#FF0000">你不具备此权限，请与管理员联系！</font></center>
<%  
  response.end
end if
%>

<%
if request("hid1")="ok" then
sql="select * from shipment where contract='"&request("contract")&"'"
set rs=conn.execute(sql)
if rs.eof=false then
%>
<script language="javascript">
alert("您输入的合同号已经存在，请重新输入！")
window.history.go(-1)
</script> 
<%
  response.end
end if
nowtrantype=request("trantype")
nowstatus=request("status")
nowinvoicestatus=request("invoicestatus")
nowcustomer=request("customer")
nowvendor=request("vendor")
nowcontract=request("contract")
nowcountry=request("country")

nowincoterm=request("incoterm")
nowdongjiancom=request("dongjiancom")
if request("dongjian")="单击选择动检许可证" then
	nowdongjian=""
else
	nowdongjian=request("dongjian")
end if
nowzidongcom=request("zidongcom")
if request("zidong")="单击选择自动许可证" then
	nowzidong=""
else
	nowzidong=request("zidong")
end if
nowplanship=request("planship")
nowplant=request("plant")
nowshipdate=request("shipdate")
nowboarddate=request("boarddate")
nowcasenumber=request("casenumber")
nowcase=request("case")
nowweishengzheng=request("weishengzheng")
nowdeliverydate=request("deliverydate")
nowladnumber=request("ladnumber")
nowshipname=request("shipname")
nownetweight=request("netweight")
nowtariff=request("tariff")
nowvat=request("vat")
nowwarranty=request("warranty")
nowclaim=request("claim")
nowapplycustom=request("applycustom")
nowpaydate=request("paydate")
nowpayment=request("payment")
nowcurrency=request("currency")
nowmemo=request("memo")
nowcargodate=request("cargodate")
nowcargodir=request("cargodir")
nowcanceldoc=request("canceldoc")


if request("trantype")="代理" then
	sql="select top 1 * from shipment where shipmentnum like '1%' order by shipmentnum desc"
else
	sql="select top 1 * from shipment where shipmentnum like '2%' order by shipmentnum desc"
end if
set rs_count=conn.execute(sql)

if rs_count.bof and rs_count.eof then
	if request("trantype")="代理" then
		nowshipment="10000001"
	else
		nowshipment="20000001"
	end if
else
	rs_count.movefirst
	nowshipment=rs_count("shipmentnum") + 1
end if	


sql="insert into shipment(shipmentnum,trantype,status,invoicestatus,vendor,contract,country,incoterm,dongjiancom,dongjian,zidongcom,zidong,planship,plant,createdate,creator) values("&nowshipment&",'"&nowtrantype&"','"&nowstatus&"','"&nowinvoicestatus&"','"&nowvendor&"','"&nowcontract&"','"&nowcountry&"','"&nowincoterm&"','"&nowdongjiancom&"','"&nowdongjian&"','"&nowzidongcom&"','"&nowzidong&"','"&nowplanship&"','"&nowplant&"',#"&now()&"#,'"&session("redboy_username")&"')"
conn.execute(sql)

for i = 1 to request("itemno").count
nowitemno=i
nowmatitem=request("material"&i)
nowcusitem=request("customer"&i)
nowspecitem=request("spec"&i)
nowcontweightitem=request("contractweight"&i)
nowactualweightitem=request("actualweight"&i)
nowpurpriceitem=request("purchaseprice"&i)
nowproducedateitem=request("produceDate"&i)
nowcasenumitem=request("casenum"&i)
nowinvamtitem=request("invamount"&i)
nowcurritem=request("invcurr"&i)
nowfinalamtitem=request("finalamount"&i)

'sql="insert into shipmentitem(shipmentnum,itemnum,material,customer,spec,contractweight,contractweightuom,purchaseprice,purchasepriceunit,productiondate,casenumber,actualnetweight,actualnetweightuom,invoiceamount,invoicecurrency,finalpayment,finalpaymentcurrency) values("&nowshipment&",'"&nowitemno&"','"&nowmatitem&"','"&nowcusitem&"','"&nowspecitem&"','"&nowcontweightitem&"','公斤','"&nowpurpriceitem&"','公斤','"&nowproducedateitem&"',"&nowcasenumitem&",'"&nowactualweightitem&"','公斤','"&nowinvamtitem&"','"&nowcurritem&"','"&nowfinalamtitem&"','"&nowcurritem&"')"
sql="insert into shipmentitem(shipmentnum,itemnum,material,customer,spec,contractweight,contractweightuom,actualnetweight,actualnetweightuom,purchaseprice,purchasepriceunit,productiondate,casenumber,invoiceamount,invoicecurrency,finalpayment,finalpaymentcurrency) values("&nowshipment&",'"&nowitemno&"','"&nowmatitem&"','"&nowcusitem&"','"&nowspecitem&"','"&nowcontweightitem&"','公斤','"&nowactualweightitem&"','公斤','"&nowpurpriceitem&"','公斤',#"&nowproducedateitem&"#,'"&nowcasenumitem&"',"&nowinvamtitem&",'"&nowcurritem&"',"&nowfinalamtitem&",'"&nowcurritem&"')"
conn.execute(sql)

next

%>
<script language="javascript">
//alert(&nowshipment&)
alert("船期表录入成功！")
window.location.href="shipment.asp"
</script>

<%
else
%>
<script language="javascript">

<% 
'二级数据保存到数组 
dim vendorCount
sql="select * from vendor order by vendorname"
set rs_vendor=conn.execute(sql)
%> 
var country = new Array(); 
var plant = new Array();
var incoterm = new Array();
//数组结构：一级根值,二级根值,二级显示值 
<% 
vendorCount = 0 
do while not rs_vendor.eof 
%> 
//set country array
country[<%=vendorCount%>] = new Array('<%=rs_vendor("vendorname")%>','<%=rs_vendor("country")%>');

plant[<%=vendorCount%>] = new Array('<%=rs_vendor("vendorname")%>','<%=rs_vendor("plant1")%>','<%=rs_vendor("plant2")%>','<%=rs_vendor("plant3")%>','<%=rs_vendor("plant4")%>');

incoterm[<%=vendorCount%>] = new Array('<%=rs_vendor("vendorname")%>','<%=rs_vendor("term1")%>','<%=rs_vendor("term2")%>');
<% 
vendorCount = vendorCount + 1 

rs_vendor.movenext 
loop 
rs_vendor.close 
%> 

<%
'保存品名数据到数组
dim materialCount
dim material()
sql="select * from material order by materialname"
set rs_mat=conn.execute(sql)
do while not rs_mat.eof 
'set material array
redim Preserve material(materialCount)
material(materialCount)=rs_mat("materialname")
materialCount = materialCount + 1 
rs_mat.movenext 
loop 
rs_mat.close 
%>   

<%
'保存客户数据到数组
dim customerCount
dim customer()
sql="select * from customer order by customername"
set rs_customer=conn.execute(sql)
do while not rs_customer.eof 
'set customer array
redim Preserve customer(customerCount)
customer(customerCount)=rs_customer("customername")
customerCount = customerCount + 1 
rs_customer.movenext 
loop 
rs_customer.close 
%>   

<%
'保存用户数据到数组
dim buyerCount, salesCount, customCount
dim buyer()
dim sales()
dim custom()
sql="select * from login order by username"
set rs_user=conn.execute(sql)
do while not rs_user.eof 
'set buyer array
if rs_user("ispurchase")="True" then
redim Preserve buyer(buyerCount)
buyer(buyerCount)=rs_user("username")
'buyer(buyerCount)=1
buyerCount = buyerCount + 1 
end if
'set sales array
if rs_user("issales")="True" then
redim Preserve sales(salesCount)
sales(salesCount)=rs_user("username")
salesCount = salesCount + 1 
end if
'set custom array
if rs_user("iscustom")="True" then
redim Preserve custom(customCount)
custom(customCount)=rs_user("username")
customCount = customCount + 1 
end if
rs_user.movenext 
loop 
rs_user.close 
%>  
      	
function check1(){
<!-- 检查合同号 -->
if (document.form1.contract.value=="")
{
alert("请输入合同号！");
return false;
}
<!-- 检查自动证 -->
<!-- 检查自动证 -->
}


function chsel(vendor){
//设置国家	
    document.form1.country.length = 0; 
    for (i=0; i<country.length; i++) 
    { 
        if (country[i][0]==vendor) 
        {document.form1.country.options[0] = new Option(country[i][1]);} 
    } 
//设置工厂    
    document.form1.plant.length = 0; 
    for (m=0; m<plant.length; m++) 
    {    	
      if (plant[m][0] == vendor) 
      {
      	for (n=1;n<5;n++)
      	{
      		if (plant[m][n]!='') {
      			document.form1.plant.options[document.form1.plant.length] = new Option(plant[m][n]);
      		}
      	} 
    	}  
    }       
//设置付款条件   
    document.form1.incoterm.length = 0; 
    for (m=0; m<incoterm.length; m++) 
    {    	
      if (incoterm[m][0] == vendor) 
      {
      	for (n=1;n<3;n++)
      	{
      		if (incoterm[m][n]!='') {
      			document.form1.incoterm.options[document.form1.incoterm.length] = new Option(incoterm[m][n]);
      		}
      	} 
    	}  
    }       
}
 
  //窗口表格增加一行
  function addNewRow(){
   var tabObj=document.getElementById("item");//获取添加数据的表格
   var rowsNum = tabObj.rows.length;  //获取当前行数
   var colsNum=tabObj.rows[0].cells.length;//获取行的列数
   var myNewRow = tabObj.insertRow(rowsNum);//插入新行.
      
   var itemObj=document.getElementsByName("itemno");//取得所有行的itemno
   var itemNo=1;
   if (itemObj.length==0) {
  	 itemNo=1;//如果没有item,给1
  	 
  }else{
  	 itemNo=parseInt(itemObj[itemObj.length-1].value) + 1; //取最大行itemno加1
  } 
   var newTdObj1=myNewRow.insertCell(0);
   newTdObj1.innerHTML="<input type='checkbox' name='chkArr'  id='chkArr' />";
   var newTdObj2=myNewRow.insertCell(1);
   newTdObj2.innerHTML="<input type='text' name='itemno' id='itemno' style='width:30px' value='"+itemNo+"' readonly='true'/>"; 
   var newTdObj3=myNewRow.insertCell(2);
	 newTdObj3.innerHTML="<select name='material"+itemNo+"' id='material"+itemNo+"' style='width:100px'>"
	    +" <% for i = 0 to materialCount-1 %>"
	 		+" <option value='<%=material(i)%>'><%=material(i)%></option>"
	 		+" <% next %> </select>";
   var newTdObj4=myNewRow.insertCell(3);
//   newTdObj4.innerHTML="<input type='text' name='customer"+itemNo+"' id='customer"+itemNo+"' />";
	 newTdObj4.innerHTML="<select name='customer"+itemNo+"' id='customer"+itemNo+"' style='width:100px'>"
	    +" <% for i = 0 to customerCount-1 %>"
	 		+" <option value='<%=customer(i)%>'><%=customer(i)%></option>"
	 		+" <% next %> </select>";   
   var newTdObj5=myNewRow.insertCell(4);
   newTdObj5.innerHTML="<input type='text' name='spec"+itemNo+"' id='spec"+itemNo+"' style='width:100px'/>";   
   var newTdObj6=myNewRow.insertCell(5);
   newTdObj6.innerHTML="<input type='text' name='contractWeight"+itemNo+"' id='contractWeight"+itemNo+"' style='width:80px' value=0 onKeyPress=\""+"javascript:CheckNum();\""+" onKeyUp=\""+"this.value=this.value.replace(/[^\\d.]/g,'')\""+">";   		 
   var newTdObj7=myNewRow.insertCell(6);
   newTdObj7.innerHTML="公斤";   
   var newTdObj8=myNewRow.insertCell(7);
   newTdObj8.innerHTML="<input type='text' name='actualWeight"+itemNo+"' id='actualWeight"+itemNo+"' style='width:80px' value=0 onKeyPress=\""+"javascript:CheckNum();\""+" onKeyUp=\""+"this.value=this.value.replace(/[^\\d.]/g,'')\""+">";   		 
   var newTdObj9=myNewRow.insertCell(8);
   newTdObj9.innerHTML="<input type='text' name='purchasePrice"+itemNo+"' id='purchasePrice"+itemNo+"' style='width:80px' value=0 onKeyPress=\""+"javascript:CheckNum();\""+" onKeyUp=\""+"this.value=this.value.replace(/[^\\d.]/g,'')\""+">";   		 
   var newTdObj10=myNewRow.insertCell(9);
   newTdObj10.innerHTML="公斤";     
   var newTdObj11=myNewRow.insertCell(10);
   newTdObj11.innerHTML="<input name='produceDate"+itemNo+"' id='produceDate"+itemNo+"' readonly style='width:80px'"
//   	+ " <img src='../images/date.gif' align='absmiddle' style='cursor:pointer;'" 
   	+ " onClick=\""+"JavaScript:window.open('../day.asp?field=produceDate"+itemNo+"&oldDate=produceDate"+itemNo+".value','','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');\"/>";     
   var newTdObj12=myNewRow.insertCell(11);   	//箱数
	 newTdObj12.innerHTML="<input name='casenum"+itemNo+"' id='casenum"+itemNo+"' style='width:80px' onKeyPress=\""+"javascript:CheckNum();\""+" onKeyUp=\""+"this.value=this.value.replace(/[^\\d.]/g,'')\""+">";   	
   var newTdObj13=myNewRow.insertCell(12);   	//发票总金额
	 newTdObj13.innerHTML="<input name='invAmount"+itemNo+"' id='invAmt"+itemNo+"' style='width:80px' value=0 onKeyPress=\""+"javascript:CheckNum();\""+" onKeyUp=\""+"this.value=this.value.replace(/[^\\d.]/g,'')\""+">";   		 
   var newTdObj14=myNewRow.insertCell(13);
   newTdObj14.innerHTML="<select name='invCurr"+itemNo+"' id='invCurr"+itemNo+"' style='width:50px'>"
	 		+" <option value='美元'>美元</option>"
	 		+" <option value='澳币'>澳币</option>"
	 		+" </select>";   
   var newTdObj15=myNewRow.insertCell(14);   	//尾款金额
	 newTdObj15.innerHTML="<input name='finalAmount"+itemNo+"' id='finalAmt"+itemNo+"' style='width:80px' value=0 onKeyPress=\""+"javascript:CheckNum();\""+" onKeyUp=\""+"this.value=this.value.replace(/[^\\d.]/g,'')\""+">";   		 

  }

//窗口表格删除一行
  function removeRow(){
   var chkObj=document.getElementsByName("chkArr");
   var tabObj=document.getElementById("item");
   for(var k=0;k<chkObj.length;k++){
    if(chkObj[k].checked){
     tabObj.deleteRow(k+1);
     k=-1;
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
      <td>&nbsp;船期表 </td>
	  <td align="right">&nbsp;</td>
    </tr>
  </table>
</td>
<td><img src="../images/r_2.gif" alt="" /></td>
</tr>
<tr>
<td></td>
<td>
	  <form name="form1">	
<table align="center" cellpadding="4" cellspacing="1" class="toptable grid" border="1">
<!--	  <form name="form1">	-->
      <tr>
        <td width="20%" align="right" height="30">一般信息：</td>
        <td width="80%" class="category">
        	类型：
					<%
					sql="select * from trantype"
					set rs_trantype=conn.execute(sql)
					%>
					<select name="trantype">
						<%do while rs_trantype.eof=false%>
						<option value="<%=rs_trantype("trantype")%>"><%=rs_trantype("trantype")%></option>
						<%
							rs_trantype.movenext
							loop
						%>
			</select>
					&nbsp;&nbsp;&nbsp;&nbsp;状态
					<%
					sql="select * from status"
					set rs_status=conn.execute(sql)
					%>
					<select name="status">
						<%do while rs_status.eof=false%>
						<option value="<%=rs_status("status")%>"><%=rs_status("status")%></option>
						<%
							rs_status.movenext
							loop
						%>
					</select>
					&nbsp;&nbsp;&nbsp;&nbsp;结算单状态
					<%
					sql="select * from invoicestatus"
					set rs_invoicestatus=conn.execute(sql)
					%>
					<select name="invoicestatus">
						<%do while rs_invoicestatus.eof=false%>
						<option value="<%=rs_invoicestatus("status")%>"><%=rs_invoicestatus("status")%></option>
						<%
							rs_invoicestatus.movenext
							loop
						%>
					</select>					
			  </td>
			</tr>
      <tr>	  
	    	<td align="right" height="30">采购：</td>
        <td class="category">		
					<select name="buyer">
	    			<% for i = 0 to buyerCount-1 %>
	 						<option value=<%=buyer(i)%>><%=buyer(i)%></option>
	 					<% next %>  			
					</select>						
					&nbsp;&nbsp;&nbsp;&nbsp;跟单	      
					<select name="custom">
	    			<% for i = 0 to customCount-1 %>
	 						<option value=<%=custom(i)%>><%=custom(i)%></option>
	 					<% next %>  			
					</select>				
					&nbsp;&nbsp;&nbsp;&nbsp;销售	      
					<select name="sales">
	    			<% for i = 0 to salesCount-1 %>
	 						<option value=<%=sales(i)%>><%=sales(i)%></option>
	 					<% next %>  			
					</select>									  		
				</td>
      </tr>			
      <tr>	  
	    	<td align="right" height="30">供应商：</td>
        <td class="category">
					<%
					sql="select * from vendor order by vendorname"
					set rs_vendor=conn.execute(sql)
					%>
					<select name="vendor" onchange="chsel(this.value)">
						<option selected="selected">---请选择---</option>	
						<%
							do while rs_vendor.eof=false
						%>
							<option value="<%=rs_vendor("vendorname")%>"><%=rs_vendor("vendorname")%></option>
						<%
							rs_vendor.movenext
							loop
							rs_vendor.close
						%>
					</select>
					&nbsp;<font color="#ff0000">*</font>
					&nbsp;&nbsp;&nbsp;&nbsp;国别				
					<select name="country">			
						<option selected="selected">---请选择---</option>		  			
					</select>						
					&nbsp;&nbsp;&nbsp;&nbsp;厂号						
					<select name="plant"> <!-- onchange="this.selectedIndex=this.defaultIndex;">-->
						  <option selected="selected">---请选择---</option>		
					</select>					
		  		&nbsp;&nbsp;&nbsp;付款方式
<!--		  		<input name="incoterm" style="width:150px">							 -->
					<select name="incoterm">			
						  <option selected="selected">---请选择---</option>										
					</select>				  				
				</td>
      </tr>
      <tr>	  
	    	<td align="right" height="30">合同号：</td>
        <td class="category">
		  		<input name="contract" value="<%=contract%>" style="width:150px" maxlength="20">
		  		&nbsp;<font color="#ff0000">*</font>
		  		&nbsp;&nbsp;&nbsp;&nbsp;品类
					<select name="materialtype">			
						  <option value="猪">猪</option>		
						  <option value="牛">牛</option>		
						  <option value="鸡">鸡</option>		
						  <option value="火鸡">火鸡</option>										
					</select>		 
		  		&nbsp;&nbsp;&nbsp;&nbsp;代理公司
					<%
					sql="select * from agent order by company"
					set rs_agent=conn.execute(sql)
					%>		  		
					<select name="agent">			
						<%
							do while rs_agent.eof=false
						%>
							<option value="<%=rs_agent("company")%>"><%=rs_agent("company")%></option>
						<%
							rs_agent.movenext
							loop
							rs_agent.close
						%>								
					</select>		  	
		  		&nbsp;&nbsp;&nbsp;&nbsp;进口批文
					<%
					sql="select * from piwen order by company"
					set rs_piwen=conn.execute(sql)
					%>		  		
					<select name="piwen">			
						<%
							do while rs_piwen.eof=false
						%>
							<option value="<%=rs_piwen("company")%>"><%=rs_piwen("company")%></option>
						<%
							rs_piwen.movenext
							loop
							rs_piwen.close
						%>								
					</select>			
					&nbsp;&nbsp;&nbsp;两证
		  		<input name="twodocumentsready" id="twodocumentsready" readonly style="width:80px">
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=twodocumentsready&oldDate='+twodocumentsready.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">										 		
				</td>
      </tr>    
      <tr>	  
	    	<td align="right" height="30">港口：</td>
        <td class="category">
					<select name="destination">			
						  <option value="上海">上海</option>		
						  <option value="大连">大连</option>		
						  <option value="天津">天津</option>		
						  <option value="待定">待定</option>										
					</select>		
					&nbsp;&nbsp;&nbsp;&nbsp;靠泊码头
					<select name="terminal">			
						  <option value="天津外代">天津外代</option>		
						  <option value="外五">外五</option>		
						  <option value="洋一">洋一</option>		
						  <option value="洋三">洋三</option>										
					</select>			
					&nbsp;&nbsp;&nbsp;&nbsp;船公司
					<select name="carrier">			
						  <option value="ANL">ANL</option>		
						  <option value="MSC">MSC</option>		
						  <option value="CMA CGM">CMA CGM</option>		
						  <option value="COSCO">COSCO</option>										
					</select>				
					&nbsp;&nbsp;&nbsp;&nbsp;船名航次
					<input name="shipname" style="width:150px" maxlength="20">							
				</td>
      </tr>               
      <tr>	  
	    	<td align="right" height="30">动检证：</td>
        <td class="category">
		  		<input name="dongjian" readonly onClick="JavaScript:window.open('../huiyuan/query_license.asp?queryform=form1&licensetype=动检证&field=dongjian&field2=dongjiancom','','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=1200,height=470,top=100,left=20');" style="cursor:hand;width:120px" value="单击选择动检许可证">
		  		<input name="dongjiancom" readonly >
				</td>
      </tr>
      <tr>	  
	    	<td align="right" height="30">自动证：</td>
        <td class="category">
		  		<input name="zidong" readonly onClick="JavaScript:window.open('../huiyuan/query_license.asp?queryform=form1&licensetype=自动证&field=zidong&field2=zidongcom','','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=1200,height=470,top=100,left=20');" style="cursor:hand;width:120px" value="单击选择自动许可证">
		  		<input name="zidongcom" readonly >
		  		&nbsp;&nbsp;&nbsp;&nbsp;自动证交批日期
		  		<input name="zidongapplydate" id="zidongapplydate" readonly style="width:80px">
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=zidongapplydate&oldDate='+zidongapplydate.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">		  				  		
		  		&nbsp;&nbsp;&nbsp;&nbsp;自动证上报日期
		  		<input name="zidongreportdate" id="zidongreportdate" readonly style="width:80px">
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=zidongreportdate&oldDate='+zidongreportdate.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">		  				  		
				</td>
      </tr>
      <tr>	  
	    	<td align="right" height="30">预保日：</td>
        <td class="category">
		  		<input name="planinsurance" id="planinsurance" style="width:80px">
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=planinsurance&oldDate='+planinsurance.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">
		  		&nbsp;&nbsp;&nbsp;&nbsp;补保日
		  		<input name="supinsurance" id="supinsurance" readonly style="width:80px">
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=supinsurance&oldDate='+supinsurance.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">
		  		&nbsp;&nbsp;&nbsp;&nbsp;保费支付日
		  		<input name="insurancepayment" id="insurancepayment" readonly style="width:80px">
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=insurancepayment&oldDate='+insurancepayment.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">		  		
		  		&nbsp;&nbsp;&nbsp;&nbsp;保单
		  		<input name="insurancenumber" style="width:120px">
				</td>
      </tr>      
      <tr>	  
	    	<td align="right" height="30">预计装船月份：</td>
        <td class="category">
		  		<input name="planship" style="width:100px">
		  		&nbsp;&nbsp;&nbsp;&nbsp;实际装船期
		  		<input name="shipdate" id="shipdate" readonly style="width:80px">
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=shipdate&oldDate='+shipdate.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">
		  		&nbsp;&nbsp;&nbsp;&nbsp;预计到港期
		  		<input name="boarddate" id="boarddate" readonly style="width:80px">
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=boarddate&oldDate='+boarddate.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">		  		
		  		&nbsp;&nbsp;&nbsp;&nbsp;客户交货期
		  		<input name="plandeliverydate" style="width:100px">
				</td>
      </tr>
      <tr>	  
	    	<td align="right" height="30">箱号：</td>
        <td class="category">
		  		<input name="case" style="width:120px" maxlength="15">
		  		&nbsp;&nbsp;&nbsp;提单号
		  		<input name="ladnumber" style="width:120px" maxlength="20">
		  		&nbsp;&nbsp;&nbsp;卫生证号
		  		<input name="weishengzheng" style="width:120px" maxlength="20">	
					&nbsp;&nbsp;&nbsp;铅封号
					<input name="locknumber" style="width:120px" maxlength="20">		
		  		&nbsp;&nbsp;&nbsp;电子信息
					<select name="einformation">			
						  <option value="有">有</option>		
						  <option value="无">无</option>										
					</select>	
		  		&nbsp;&nbsp;&nbsp;清真证明
					<select name="MuslimCertification">			
						  <option value="有">有</option>		
						  <option value="无">无</option>										
					</select>											  		
				</td>
      </tr>
      <tr>	  
	    	<td align="right" height="30">预付款日期：</td>
        <td class="category">
		  		<input name="paydate" id="paydate" readonly style="width:80px">
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=paydate&oldDate='+paydate.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">
		  		&nbsp;&nbsp;&nbsp;预付款金额
		  		<input name="payment" style="width:100px" maxlength="15" value=0 onKeyPress="javascript:CheckNum();" onKeyUp="this.value=this.value.replace(/[^\d.]/g,'')">
		  		&nbsp;&nbsp;&nbsp;币种
					<%
					sql="select * from trancurrency"
					set rs_currency=conn.execute(sql)
					%>
					<select name="currency">
						<%
							do while rs_currency.eof=false
						%>
							<option value="<%=rs_currency("currency")%>"><%=rs_currency("currency")%></option>
						<%
							rs_currency.movenext
							loop
						%>
					</select>
				&nbsp;&nbsp;&nbsp;订金收到日期
				<input name="downpaymentdate" id="downpaymentdate" readonly style="width:80px">	
				<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=downpaymentdate&oldDate='+downpaymentdate.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">
				&nbsp;&nbsp;&nbsp;成交条款
					<select name="tradingterm">
							<option value="CIF">CIF</option>
							<option value="CFR">CFR</option>
							<option value="CIP">CIP</option>
							<option value="CNF">CNF</option>
							<option value="CPT">CPT</option>
					</select>
				</td>
      </tr> 
      <tr>	  
	    	<td align="right" height="30">尾款支付日期：</td>
        <td class="category">
		  		<input name="finalpaymentdate" id="finalpaymentdate" readonly style="width:80px">
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=finalpaymentdate&oldDate='+finalpaymentdate.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">
		  		&nbsp;&nbsp;&nbsp;免箱期
		  		<input name="freestayperiod" style="width:100px">
		  		&nbsp;&nbsp;&nbsp;卫生证版本
					<input name="weishengzhengversion" style="width:100px">
					&nbsp;&nbsp;&nbsp;放行日期
		  		<input name="passdate" id="passdate" readonly style="width:80px">
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=passdate&oldDate='+passdate.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">
				</td>
      </tr>       
      <tr>	  
	    	<td align="right" height="30">到单日期：</td>
        <td class="category">
		  		<input name="documentarrival" id="documentarrival" readonly style="width:80px">
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=documentarrival&oldDate='+documentarrival.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">
		  		&nbsp;&nbsp;&nbsp;预计赎单日期
		  		<input name="planretirebill" id="planretirebill" readonly style="width:80px">
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=planretirebill&oldDate='+planretirebill.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">		  		
		  		&nbsp;&nbsp;&nbsp;我司赎单日期
		  		<input name="internalretirebill" id="internalretirebill" readonly style="width:80px">
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=internalretirebill&oldDate='+internalretirebill.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">		  		
		  		&nbsp;&nbsp;&nbsp;代理赎单日期
		  		<input name="externalretirebill" id="externalretirebill" readonly style="width:80px">
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=externalretirebill&oldDate='+externalretirebill.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">		  		
				</td>
      </tr>       
      <tr>	  
	    	<td align="right" height="30">送货日期：</td>
        <td class="category">
		  		<input name="cargodate" id="cargodate" readonly style="width:80px">
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=cargodate&oldDate='+cargodate.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">
		  		&nbsp;&nbsp;&nbsp;送货方向
		  		<input name="cargodir" style="width:150px">
		  		&nbsp;&nbsp;&nbsp;交单日期
		  		<input name="deliverydate" id="deliverydate" readonly style="width:80px">
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=deliverydate&oldDate='+deliverydate.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">		  		
				</td>
      </tr> 
      <tr>
        <td align="right" height="30" valign="top">单证问题：</td>
        <td class="category" valign="top">
        	<textarea name="documentissue" cols="30" rows="3"></textarea>&nbsp;&nbsp;&nbsp;
        	<label for="documentcheck" style="vertical-align:top"> 审单情况 </label>
        	<textarea name="documentcheck" id="documentcheck" cols="30" rows="3"></textarea>&nbsp;&nbsp;&nbsp;
        	<label for="documentdelivery" style="vertical-align:top"> 寄单情况 </label>
        	<textarea name="documentdelivery" id="documentdelivery" cols="30" rows="3"></textarea>&nbsp;&nbsp;&nbsp;
        	<label for="clearcustom" style="vertical-align:top">清关进程</label>
        	<textarea name="clearcustom" id="clearcustom" cols="30" rows="3"></textarea>            	   	
        </td>
      </tr>	 
      <tr>
        <td align="right" height="30" valign="top">保证金情况：</td>
        <td class="category" valign="top">
        	<textarea name="warrantyinformation" cols="30" rows="3"></textarea>&nbsp;&nbsp;&nbsp;
        	<label for="warrantystatus" style="vertical-align:top">保证金进度</label>
					<select name="warrantystatus" style="vertical-align:top">
	 						<option value="审核中">审核中</option>	
	 						<option value="已退保">已退保</option>	
	 						<option value="已到账">已到账</option>	
					</select>			
					<label for="warranty" style="vertical-align:top">&nbsp;&nbsp;&nbsp;保证金金额</label>	
					<input name="warranty" style="vertical-align:top;width:100px" maxlength="15" value=0 onKeyPress="javascript:CheckNum();" onKeyUp="this.value=this.value.replace(/[^\d.]/g,'')">
					<label for="returnwarranty" style="vertical-align:top">&nbsp;&nbsp;&nbsp;退保证金金额</label>	
					<input name="returnwarranty" style="vertical-align:top;width:100px" maxlength="15" value=0 onKeyPress="javascript:CheckNum();" onKeyUp="this.value=this.value.replace(/[^\d.]/g,'')">
        </td>
      </tr>	        
      <tr>
        <td align="right" height="30" valign="top">少货情况：</td>
        <td class="category" valign="top">
        	<textarea name="stockmiss" cols="30" rows="3"></textarea>&nbsp;&nbsp;&nbsp;
        	<label for="claiminformation" style="vertical-align:top">索赔情况</label>
        	<textarea name="claiminformation" id="claiminformation" cols="30" rows="3"></textarea>&nbsp;&nbsp;&nbsp;
					<label for="claimstatus" style="vertical-align:top">索赔状态</label>
					<select name="claimstatus" style="vertical-align:top">
							<option value="索赔中">索赔中</option>
							<option value="需索赔">需索赔</option>
							<option value="不需索赔">不需索赔</option>
					</select>     
					<label for="claimprocessor" style="vertical-align:top">索赔负责人</label>
					<select name="custom" style="vertical-align:top">
	    			<% for i = 0 to customCount-1 %>
	 						<option value=<%=custom(i)%>><%=custom(i)%></option>
	 					<% next %>  			
					</select>								
        </td>
      </tr>	       
      <tr>
		  <input type="submit" value=" 确认录入 " onClick="return check1();" class="button">&nbsp;&nbsp;&nbsp;
		  <input type="hidden" name="hid1" value="ok">
		  <input type="reset" value=" 重新填写 " class="button">
      </tr>    
</table>  
<table align="center" cellpadding="4" cellspacing="1" class="toptable grid" border="1">
<!--	  <form name="formitem" action="" method="post">-->
      <tr>
					<input type="button" onclick="addNewRow()" value="增加一行" class="button"/>&nbsp;&nbsp;&nbsp;
					<input type="button" onclick="removeRow()" value="删除一行" class="button"/>
      </tr>	      
      <tr>
					<table align="center" cellpadding="4" cellspacing="1" class="toptable grid" border="1" id="item">
					  <tr>
					  	<td width="20"></td>			
			        <td width="30" align="center" >项目号</td>
			        <td width="100" align="center" >品名</td>
			        <td width="100" align="center" >客户</td> 	
			        <td width=100" align="center" >规格</td> 		
			        <td width=80" align="center" >合同重量</td> 			
			        <td width="40" align="center" >重量单位</td> 		
			        <td width="80" align="center" >实际净重</td> 				    
			        <td width="80" align="center" >采购价格</td> 			
			        <td width="40" align="center" >价格单位</td> 				        	            		        	        		        
			        <td width="50" align="center" >生产日期</td> 		
			        <td width="50" align="center" >箱数</td> 		
			        <td width="100" align="center" >发票总金额</td> 		
			        <td width="50" align="center" >币种</td> 		
			        <td width="100" align="center" >尾款金额</td> 					        			        			        			        			        
			      </tr>			
			    </table>			  	
      </tr>	          

<!--  </form>-->
</table>
</form>
</td>
<td></td>
</tr> 
</table>

<%end if%>
</body>
</html>