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
<title><%=dianming%> - 修改船期表</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script type="text/javascript" src="../js/jquery-3.2.1.js"></script>
<!--<script type="text/javascript" src="https://code.jquery.com/jquery-1.12.4.js"></script>-->
<script type="text/javascript" src="../js/jquery-ui.js"></script>
<!--<script type="text/javascript" src="../js/jquery.dataTables.min.js"></script>-->

<!--
<script type="text/javascript" src="https://cdn.datatables.net/v/dt/jq-3.2.1/jq-3.2.1/moment-2.18.1/dt-1.10.16/b-1.4.2/sl-1.2.3/datatables.min.js"></script>
<script type="text/javascript" src="../js/dataTables.editor.min.js"></script>
-->

<link href="../style/jquery-ui.css" rel="stylesheet" type="text/css">
<link href="../style/style.css" rel="stylesheet" type="text/css">

<!--X-Editable test -->
<script src="../js/test_demo/assets/mockjax/jquery.mockjax.js"></script>
<!-- momentjs --> 
<script src="../js/test_demo/assets/momentjs/moment.min.js"></script> 
<!-- bootstrap 3 -->
<link href="../js/test_demo/assets/bootstrap300/css/bootstrap.css" rel="stylesheet">
<script src="../js/test_demo/assets/bootstrap300/js/bootstrap.js"></script>
<!-- x-editable (bootstrap 3) -->
<link href="../js/test_demo/assets/x-editable/bootstrap3-editable/css/bootstrap-editable.css" rel="stylesheet">
<script src="../js/test_demo/assets/x-editable/bootstrap3-editable/js/bootstrap-editable.js"></script>      
<!-- shipment input -->
<link href="../js/test_demo/shipname.css" rel="stylesheet">
<script src="../js/test_demo/shipname.js"></script>     
<script src="../js/test_demo/assets/demo-mock.js"></script> 
<script src="../js/test_demo/assets/demo.js"></script>  


<!--
<link href="../style/jquery.dataTables.min.css" rel="stylesheet" type="text/css">
<link href="../style/editor.dataTables.min.css" rel="stylesheet" type="text/css">
-->
<style>
body {
	background-color:#FFFFFF;
}
</style>
</HEAD>

<BODY>
<%
if fla1="0" and fla2="0" and fla3="0" and fla4="0" then
%>
<br><center><img src="../images/note.gif" align="absmiddle">&nbsp;<font color="#FF0000">你不具备此权限，请与管理员联系！</font></center>
<%  
  response.end
end if
%>

<%
  sql="select * from locktable where tablename='shipment' and combinedkey='"&request("shipment")&"'"
  set rs_lock=conn.execute(sql)
  if rs_lock.eof = false then
    if rs_lock("username")<>session("redboy_username") then
%>
    <script language="javascript">
    alert("用户<%=rs_lock("username")%>正在编辑该记录！请稍后再试！");
    window.location.href="shipment.asp";
    </script> 
<%end if
else
    sql="insert into locktable(tablename,combinedkey,status,username,locktime) values('shipment','"&request("shipment")&"','E','"&session("redboy_username")&"',#"&now()&"#)"  
    conn.execute(sql)
end if
%>

<script>
  $(document).ready(function(){
//日期控件。要支持动态创建的控件
    $("body").delegate(".datepicker","focusin",function(){
      $(this).datepicker();
    });

  });
</script>

<%
if request("hid1")="ok" then
sql="select * from shipment where shipmentnum="&request("shipment")
set rs=server.createobject("ADODB.RecordSet")
rs.open sql,conn,1,3

rs("trantype")=request("trantype")
rs("status")=request("status")
rs("invoicestatus")=request("invoicestatus")
rs("vendor")=request("vendor")
rs("contract")=request("contract")
rs("country")=request("country")
rs("sales")=request("sales")
rs("incoterm")=request("incoterm")
rs("dongjiancom")=request("dongjiancom")
if request("dongjian")="单击选择动检证" then
  rs("dongjian")=""
else
  rs("dongjian")=request("dongjian")
end if
rs("zidongcom")=request("zidongcom")
if request("zidong")="单击选择自动证" then
  rs("zidong")=""
else
  rs("zidong")=request("zidong")
end if
rs("planship")=request("planship")
rs("plant")=request("plant")
if request("shipdate")<>"" then
  rs("shipdate")=request("shipdate")
end if
if request("boarddate")<>"" then
  rs("boarddate")=request("boarddate")
end if

rs("case")=request("case")
rs("weishengzheng")=request("weishengzheng")
if request("deliverydate")<>"" then
  rs("deliverydate")=request("deliverydate")
end if
rs("ladnumber")=request("ladnumber")
rs("shipname")=request("shipname")

rs("warranty")=request("warranty")
rs("claiminformation")=request("claiminformation")
rs("applycustom")=request("applycustom")
if request("paydate")<>"" then
  rs("paydate")=request("paydate")
end if
rs("prepayment")=request("prepayment")
rs("trancurrency")=request("currency")
rs("memo")=request("memo")
if request("cargodate")<>"" then
  rs("cargodate")=request("cargodate")
end if
rs("cargodir")=request("cargodir")
rs("canceldoc")=request("canceldoc")
rs("changedate")=now
rs("changer")=session("redboy_username")
rs.update
rs.close

'先删除已有的item'
sql="delete from shipmentitem where shipmentnum="&request("shipment")
conn.execute(sql)

'再插入item'
for i = 1 to request("itemno").count
nowitemno=i
nowmatitem=request("material"&i)
nowcusitem=request("customer"&i)
nowspecitem=request("spec"&i)
if request("contractweight"&i)<>"" then
  nowcontweightitem=request("contractweight"&i)
else
  nowcontweightitem=0
end if
if request("actualweight"&i)<>"" then
  nowactualweightitem=request("actualweight"&i)
else
  nowactualweightitem=0
end if
if request("purchaseprice"&i)<>"" then
  nowpurpriceitem=request("purchaseprice"&i)
else
  nowpurpriceitem=0
end if
if request("produceDate"&i)<>"" then
  nowproducedateitem=request("produceDate"&i)
else
  nowproducedateitem=DateSerial(1900, 1, 1)
end if
if request("casenum"&i)<>"" then
  nowcasenumitem=request("casenum"&i)
else
  nowcasenumitem=0
end if
if request("invamount"&i)<>"" then
  nowinvamtitem=request("invamount"&i)
else
  nowinvamtitem=0
end if
nowcurritem=request("invcurr"&i)
if request("finalamount"&i)<>"" then
  nowfinalamtitem=request("finalamount"&i)
else
  nowfinalamtitem=0
end if

sql="insert into shipmentitem(shipmentnum,itemnum,material,customer,spec,contractweight,contractweightuom,actualnetweight,actualnetweightuom,purchaseprice,purchasepriceunit,productiondate,casenumber,invoiceamount,invoicecurrency,finalpayment,finalpaymentcurrency) values("&request("shipment")&",'"&nowitemno&"','"&nowmatitem&"','"&nowcusitem&"','"&nowspecitem&"','"&nowcontweightitem&"','公斤','"&nowactualweightitem&"','公斤','"&nowpurpriceitem&"','公斤',#"&nowproducedateitem&"#,'"&nowcasenumitem&"',"&nowinvamtitem&",'"&nowcurritem&"',"&nowfinalamtitem&",'"&nowcurritem&"')"

'sql="insert into shipmentitem(shipmentnum,itemnum,material,customer,spec,contractweight,contractweightuom,actualnetweight,actualnetweightuom) values("&request("shipment")&","&nowitemno&",'"&nowmatitem&"','"&nowcusitem&"','"&nowspecitem&"',"&nowcontweightitem&",'公斤',"&nowactualweightitem&",'公斤')"
conn.execute(sql)

next

'删除逻辑锁'
sql="delete from locktable where tablename='shipment' and combinedkey='"&request("shipment")&"'"
conn.execute(sql)

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
   newTdObj1.align="center";
   var newTdObj2=myNewRow.insertCell(1);
   newTdObj2.innerHTML="<input type='text' name='itemno' id='itemno' style='align:center;width:30px' value='"+itemNo+"' readonly='true'/>"; 
   newTdObj2.align="center";
   var newTdObj3=myNewRow.insertCell(2);
	 newTdObj3.innerHTML="<select name='material"+itemNo+"' id='material"+itemNo+"' style='width:100px'>"
	    +" <% for i = 0 to materialCount-1 %>"
	 		+" <option value='<%=material(i)%>'><%=material(i)%></option>"
	 		+" <% next %> </select>";
   newTdObj3.align="center";
   var newTdObj4=myNewRow.insertCell(3);
//   newTdObj4.innerHTML="<input type='text' name='customer"+itemNo+"' id='customer"+itemNo+"' />";
	 newTdObj4.innerHTML="<select name='customer"+itemNo+"' id='customer"+itemNo+"' style='width:100px' align='middle'>"
	    +" <% for i = 0 to customerCount-1 %>"
	 		+" <option value='<%=customer(i)%>'><%=customer(i)%></option>"
	 		+" <% next %> </select>";   
   newTdObj4.align="center";
   var newTdObj5=myNewRow.insertCell(4);
   newTdObj5.innerHTML="<input type='text' name='spec"+itemNo+"' id='spec"+itemNo+"' style='width:100px' align='center'/>";   
   newTdObj5.align="center";
   var newTdObj6=myNewRow.insertCell(5);
   newTdObj6.innerHTML="<input type='text' name='contractWeight"+itemNo+"' id='contractWeight"+itemNo+"' style='width:80px' align='center' value=0 onKeyPress=\""+"javascript:CheckNum();\""+" onKeyUp=\""+"this.value=this.value.replace(/[^\\d.]/g,'')\""+">";  
   newTdObj6.align="center"; 		 
   var newTdObj7=myNewRow.insertCell(6);
   newTdObj7.innerHTML="公斤";   
   newTdObj7.align="center";
   var newTdObj8=myNewRow.insertCell(7);
   newTdObj8.innerHTML="<input type='text' name='actualWeight"+itemNo+"' id='actualWeight"+itemNo+"' style='width:80px'  align='center' value=0 onKeyPress=\""+"javascript:CheckNum();\""+" onKeyUp=\""+"this.value=this.value.replace(/[^\\d.]/g,'')\""+">";   
   newTdObj8.align="center";		 
   var newTdObj9=myNewRow.insertCell(8);
   newTdObj9.innerHTML="<input type='text' name='purchasePrice"+itemNo+"' id='purchasePrice"+itemNo+"' style='width:80px' align='center' value=0 onKeyPress=\""+"javascript:CheckNum();\""+" onKeyUp=\""+"this.value=this.value.replace(/[^\\d.]/g,'')\""+">";   	
   newTdObj9.align="center";	 
   var newTdObj10=myNewRow.insertCell(9);
   newTdObj10.innerHTML="公斤";     
   newTdObj10.align="center";
   var newTdObj11=myNewRow.insertCell(10);
   newTdObj11.innerHTML="<input name='produceDate"+itemNo+"' id='produceDate"+itemNo+"' style='width:80px' class='datepicker'>"
//   	+ " <img src='../images/date.gif' align='absmiddle' style='cursor:pointer;'" 
//   	+ " onClick=\""+"JavaScript:window.open('../day.asp?form=formitem&field=produceDate"+itemNo+"&oldDate=produceDate"+itemNo+".value','','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');\"/>";   
   newTdObj11.align="center";  
   var newTdObj12=myNewRow.insertCell(11);   	//箱数
	 newTdObj12.innerHTML="<input name='casenum"+itemNo+"' id='casenum"+itemNo+"' style='width:80px' align='center' onKeyPress=\""+"javascript:CheckNum();\""+" onKeyUp=\""+"this.value=this.value.replace(/[^\\d.]/g,'')\""+">";   
   newTdObj12.align="center";	
   var newTdObj13=myNewRow.insertCell(12);   	//发票总金额
	 newTdObj13.innerHTML="<input name='invAmount"+itemNo+"' id='invAmt"+itemNo+"' style='width:80px' align='center' value=0 onKeyPress=\""+"javascript:CheckNum();\""+" onKeyUp=\""+"this.value=this.value.replace(/[^\\d.]/g,'')\""+">";   	
   newTdObj13.align="center";	 
   var newTdObj14=myNewRow.insertCell(13);
   newTdObj14.innerHTML="<select name='invCurr"+itemNo+"' id='invCurr"+itemNo+"' style='width:50px' align='center'>"
	 		+" <option value='美元'>美元</option>"
	 		+" <option value='澳币'>澳币</option>"
	 		+" </select>";   
   newTdObj14.align="center";
   var newTdObj15=myNewRow.insertCell(14);   	//尾款金额
	 newTdObj15.innerHTML="<input name='finalAmount"+itemNo+"' id='finalAmt"+itemNo+"' style='width:80px' value=0 align='center' onKeyPress=\""+"javascript:CheckNum();\""+" onKeyUp=\""+"this.value=this.value.replace(/[^\\d.]/g,'')\""+">";   
   newTdObj15.align="center";		 

  }

//窗口表格删除一行
  function removeRow(){
   var chkObj=document.getElementsByName("chkArr");
   var tabObj=document.getElementById("item");
   var shipment=document.getElementById("shipment");
   var itemno=document.getElementsByName("itemno");
   var existing=0;

   for(var k=0;k<chkObj.length;k++){
    if(chkObj[k].checked){
//检查是否已经有结算单或入库记录。如果有，则不能删除
	<%    
	sql="select * from salescontract where refshipment="&request("shipment")
	set rs_contract=conn.execute(sql)
  	if rs_contract.eof then
	%>	
     tabObj.deleteRow(k+1);
     k=-1;
     break;
    <%else
    	do while not rs_contract.eof 
    %>
		if (itemno[k].value==<%=rs_contract("refitem")%>){
			existing=1;
		}	
    <%
    	rs_contract.movenext
    	loop
      end if
    %>
    if (existing==1) {
    	alert("已经有结算单，不能删除该项目！")
    }else{
     	tabObj.deleteRow(k+1);
     k=-1;    
    } 	   
    }
   }
  }

</script>

<%
sql="select * from shipment where shipmentnum="&request("shipment")
'set rs=conn.execute(sql)
  set rs =server.createobject("ADODB.RecordSet")	
  rs.open sql,conn,1,2
%>

<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../images/r_1.gif" alt="" /></td>
<td width="100%" background="../images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td width="10%" align="left">&nbsp;修改船期表资料(新)  ---  </td>
	  <td align="left"><%=rs("shipmentnum")%></td>
    </tr>
  </table>
</td>
<td><img src="../images/r_2.gif" alt="" /></td>
</tr>
<tr>
<td></td>
<td>
	  <form name="form1">	
<table align="center" cellpadding="4" cellspacing="1" class="toptable grid" border="1" id="maintable">
<!--	  <form name="form1">	-->
      <tbody>
      <tr>
        <td width="20%" align="right" height="30">船期表号码：</td>
		<td width="80%" class="category"><input name="shipment" id="shipment" value=<%=rs("shipmentnum")%> readonly></td>
        
      </tr>
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
						<option value="<%=rs_trantype("trantype")%>" <%if rs_trantype("trantype")=rs("trantype") then%>selected="selected"<%end if%>><%=rs_trantype("trantype")%></option>
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
						<option value="<%=rs_status("status")%>" <%if rs_status("status")=rs("status") then%>selected="selected"<%end if%>><%=rs_status("status")%></option>
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
						<option value="<%=rs_invoicestatus("status")%>" <%if rs_invoicestatus("status")=rs("invoicestatus") then%>selected="selected"<%end if%>><%=rs_invoicestatus("status")%></option>
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
	 						<option value="<%=buyer(i)%>" <%if buyer(i)=rs("buyer") then%>selected="selected"<%end if%>><%=buyer(i)%></option>
	 					<% next %>  			
					</select>						
					&nbsp;&nbsp;&nbsp;&nbsp;跟单	      
					<select name="custom">
	    			<% for i = 0 to customCount-1 %>
	 						<option value="<%=custom(i)%>" <%if custom(i)=rs("handler") then%>selected="selected"<%end if%>><%=custom(i)%></option>
	 					<% next %>  			
					</select>				
					&nbsp;&nbsp;&nbsp;&nbsp;销售	      
					<select name="sales">
	    			<% for i = 0 to salesCount-1 %>
	 						<option value="<%=sales(i)%>" <%if sales(i)=rs("sales") then%>selected="selected"<%end if%>><%=sales(i)%></option>
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
							<option value="<%=rs_vendor("vendorname")%>" <%if rs_vendor("vendorname")=rs("vendor") then%>selected="selected"<%end if%>><%=rs_vendor("vendorname")%></option>
						<%
							rs_vendor.movenext
							loop
							rs_vendor.close
						%>
					</select>
					&nbsp;<font color="#ff0000">*</font>
					&nbsp;&nbsp;&nbsp;&nbsp;国别				
					<select name="country">			
						<option selected="selected"><%=rs("country")%></option>		  			
					</select>						
					&nbsp;&nbsp;&nbsp;&nbsp;厂号						
					<select name="plant"> <!-- onchange="this.selectedIndex=this.defaultIndex;">-->
						  <option selected="selected"><%=rs("plant")%></option>		
					</select>					
		  		&nbsp;&nbsp;&nbsp;付款方式
<!--		  		<input name="incoterm" style="width:150px">							 -->
					<select name="incoterm">			
						  <option selected="selected"><%=rs("incoterm")%></option>										
					</select>				  				
				</td>
      </tr>
      <tr>	  
	    	<td align="right" height="30">合同号：</td>
        <td class="category">
		  		<input name="contract" value=<%=rs("contract")%> style="width:150px" maxlength="20">
		  		&nbsp;<font color="#ff0000">*</font>
		  		&nbsp;&nbsp;&nbsp;&nbsp;品类
					<%
					sql="select * from materialtype"
					set rs_mattype=conn.execute(sql)
					%>		  		
					<select name="materialtype">			
						<%
							do while rs_mattype.eof=false
						%>
							<option value=<%=rs_mattype("materialtype")%> <%if rs_mattype("materialtype")=rs("materialtype") then%>selected="selected"<%end if%>><%=rs_mattype("materialtype")%></option>
						<%
							rs_mattype.movenext
							loop
							rs_mattype.close
						%>		 
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
							<option value="<%=rs_agent("company")%>" <%if rs_agent("company")=rs("agent") then%>selected="selected"<%end if%>><%=rs_agent("company")%></option>
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
							<option value="<%=rs_piwen("company")%>" <%if rs_piwen("company")=rs("piwen") then%>selected="selected"<%end if%>><%=rs_piwen("company")%></option>
						<%
							rs_piwen.movenext
							loop
							rs_piwen.close
						%>								
					</select>			
					&nbsp;&nbsp;&nbsp;两证
		  		<input name="twodocumentsready" id="twodocumentsready" value="<%=rs("twodocumentready")%>" class="datepicker"  style="width:80px">										 		
				</td>
      </tr>    
      <tr>	  
	    	<td align="right" height="30">港口：</td>
        <td class="category">
					<%
					sql="select * from port"
					set rs_port=conn.execute(sql)
					%>	        
					<select name="destination">			
						<%
							do while rs_port.eof=false
						%>
							<option value=<%=rs_port("port")%> <%if rs_port("port")=rs("destination") then%>selected="selected"<%end if%>><%=rs_port("port")%></option>
						<%
							rs_port.movenext
							loop
							rs_port.close
						%>						
<!--						  <option value="上海">上海</option>		
						  <option value="大连">大连</option>		
						  <option value="天津">天津</option>		
						  <option value="待定">待定</option>				-->						
					</select>		
					&nbsp;&nbsp;&nbsp;&nbsp;靠泊码头
					<%
					sql="select * from terminal"
					set rs_terminal=conn.execute(sql)
					%>						
					<select name="terminal">		
						<%
							do while rs_terminal.eof=false
						%>
							<option value=<%=rs_terminal("terminal")%> <%if rs_terminal("terminal")=rs("terminal") then%>selected="selected"<%end if%>><%=rs_terminal("terminal")%></option>
						<%
							rs_terminal.movenext
							loop
							rs_terminal.close
						%>	
<!--						  <option value="天津外代">天津外代</option>		
						  <option value="外五">外五</option>		
						  <option value="洋一">洋一</option>		
						  <option value="洋三">洋三</option>				-->						
					</select>			
					&nbsp;&nbsp;&nbsp;&nbsp;船公司
					<%
					sql="select * from carrier"
					set rs_carrier=conn.execute(sql)
					%>							
					<select name="carrier">			
						<%
							do while rs_carrier.eof=false
						%>
							<option value=<%=rs_carrier("carrier")%> <%if rs_carrier("carrier")=rs("carrier") then%>selected="selected"<%end if%>><%=rs_carrier("carrier")%></option>
						<%
							rs_carrier.movenext 
							loop
							rs_carrier.close
						%>						
<!--						  <option value="ANL">ANL</option>		
						  <option value="MSC">MSC</option>		
						  <option value="CMA CGM">CMA CGM</option>		
						  <option value="COSCO">COSCO</option>			-->							
					</select>				
					&nbsp;&nbsp;&nbsp;&nbsp;船名航次
					<!--<input name="shipname" id="shipname" style="width:150px" maxlength="20" >		-->
					<input id="shipname" name="shipname" data-type="shipname" data-title="Please, fill shipname" value=<%=rs("shipname")%>>			
					<input type="hidden" id="shipcomment" value=<%=rs("shipcomment")%>>		
				</td>
      </tr>               
      <tr>	  
	    	<td align="right" height="30">动检证：</td>
        <td class="category">
		  		<input name="dongjian" readonly onClick="JavaScript:window.open('../huiyuan/query_license.asp?queryform=form1&licensetype=动检证&field=dongjian&field2=dongjiancom','','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=1200,height=470,top=100,left=20');" style="cursor:hand;width:120px" value=<%=rs("dongjian")%>>
		  		<input name="dongjiancom" readonly value=<%=rs("dongjiancom")%>>
				</td>
      </tr>
      <tr>	  
	    	<td align="right" height="30">自动证：</td>
        <td class="category">
		  		<input name="zidong" readonly onClick="JavaScript:window.open('../huiyuan/query_license.asp?queryform=form1&licensetype=自动证&field=zidong&field2=zidongcom','','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=1200,height=470,top=100,left=20');" style="cursor:hand;width:120px" value=<%=rs("zidong")%>>
		  		<input name="zidongcom" readonly value=<%=rs("zidongcom")%>>
		  		&nbsp;&nbsp;&nbsp;&nbsp;自动证交批日期
		  		<input name="zidongapplydate" id="zidongapplydate" class="datepicker"  style="width:80px" value=<%=rs("zidongapplydate")%>>	  				  		
		  		&nbsp;&nbsp;&nbsp;&nbsp;自动证上报日期
		  		<input name="zidongreportdate" id="zidongreportdate" class="datepicker"  style="width:80px" value=<%=rs("zidongreportdate")%>>	  				  		
				</td>
      </tr>
      <tr>	  
	    	<td align="right" height="30">预保日：</td>
        <td class="category">
		  		<input name="planinsurance" id="planinsurance" class="datepicker" style="width:80px" value=<%=rs("planinsurance")%>>
		  		&nbsp;&nbsp;&nbsp;&nbsp;补保日
		  		<input name="supinsurance" id="supinsurance" class="datepicker"  style="width:80px" value=<%=rs("supinsurance")%>>
		  		&nbsp;&nbsp;&nbsp;&nbsp;保费支付日
		  		<input name="insurancepayment" id="insurancepayment" class="datepicker" style="width:80px" value=<%=rs("insurancepayment")%>>	  		
		  		&nbsp;&nbsp;&nbsp;&nbsp;保单
		  		<input name="insurancenumber" style="width:120px" value=<%=rs("insurancenumber")%>>
				</td>
      </tr>      
      <tr>	  
	    	<td align="right" height="30">预计装船月份：</td>
        <td class="category">
		  		<input name="planship" style="width:100px">
		  		&nbsp;&nbsp;&nbsp;&nbsp;实际装船期
		  		<input name="shipdate" id="shipdate" class="datepicker" style="width:80px" value=<%=rs("shipdate")%>>
		  		&nbsp;&nbsp;&nbsp;&nbsp;预计到港期
		  		<input name="boarddate" id="boarddate" class="datepicker" style="width:80px" value=<%=rs("boarddate")%>>  		
		  		&nbsp;&nbsp;&nbsp;&nbsp;客户交货期
		  		<input name="plandeliverydate" style="width:100px" value=<%=rs("plandeliverydate")%>>
				</td>
      </tr>
      <tr>	  
	    	<td align="right" height="30">箱号：</td>
        <td class="category">
		  		<input name="case" style="width:120px" maxlength="15" value=<%=rs("case")%>>
		  		&nbsp;&nbsp;&nbsp;提单号
		  		<input name="ladnumber" style="width:120px" maxlength="20" value=<%=rs("ladnumber")%>>
		  		&nbsp;&nbsp;&nbsp;卫生证号
		  		<input name="weishengzheng" style="width:120px" maxlength="20" value=<%=rs("weishengzheng")%>>	
					&nbsp;&nbsp;&nbsp;铅封号
					<input name="locknumber" style="width:120px" maxlength="20" value=<%=rs("locknumber")%>>		
		  		&nbsp;&nbsp;&nbsp;电子信息
					<select name="einformation">			
						  <option value="有" <%if rs("einformation")="True" then%>selected="selected"<%end if%>>有</option>		
						  <option value="无" <%if rs("einformation")="False" then%>selected="selected"<%end if%>>无</option>										
					</select>	
		  		&nbsp;&nbsp;&nbsp;清真证明
					<select name="MuslimCertification">			
						  <option value="有" <%if rs("MuslimCertification")="True" then%>selected="selected"<%end if%>>有</option>		
						  <option value="无" <%if rs("MuslimCertification")="False" then%>selected="selected"<%end if%>>无</option>										
					</select>									 		  		
				</td>
      </tr>
      <tr>	  
	    	<td align="right" height="30">预付款日期：</td>
        <td class="category">
		  		<input name="paydate" id="paydate" class="datepicker" style="width:80px" value=<%=rs("prepaydate")%>>
		  		&nbsp;&nbsp;&nbsp;预付款金额
		  		<input name="prepayment" style="width:100px" maxlength="15" value=<%=rs("prepayment")%> onKeyPress="javascript:CheckNum();" onKeyUp="this.value=this.value.replace(/[^\d.]/g,'')">
		  		&nbsp;&nbsp;&nbsp;币种
					<%
					sql="select * from trancurrency"
					set rs_currency=conn.execute(sql)
					%>
					<select name="currency">
						<%
							do while rs_currency.eof=false
						%>
							<option value="<%=rs_currency("currency")%>" <%if rs("trancurrency")=rs_currency("currency") then%>selected="selected"<%end if%>><%=rs_currency("currency")%></option>
						<%
							rs_currency.movenext
							loop
						%>
					</select>
				&nbsp;&nbsp;&nbsp;订金收到日期
				<input name="downpaymentdate" id="downpaymentdate" class="datepicker" style="width:80px" value=<%=rs("downpaymentreceiptdate")%>>	
				&nbsp;&nbsp;&nbsp;成交条款
					<%
					sql="select * from tradingterm"
					set rs_term=conn.execute(sql)
					%>
					<select name="tradingterm">
						<%
							do while rs_term.eof=false
						%>
							<option value="<%=rs_term("tradingterm")%>" <%if rs("tradingterm")=rs_term("tradingterm") then%>selected="selected"<%end if%>><%=rs_term("tradingterm")%></option>
						<%
							rs_term.movenext
							loop
						%>
					</select>		
<!--							
					<select name="tradingterm">
							<option value="CIF">CIF</option>
							<option value="CFR">CFR</option>
							<option value="CIP">CIP</option>
							<option value="CNF">CNF</option>
							<option value="CPT">CPT</option>
					</select> -->
				</td>
      </tr> 
      <tr>	  
	    	<td align="right" height="30">尾款支付日期：</td>
        <td class="category">
		  		<input name="finalpaymentdate" id="finalpaymentdate" class="datepicker" style="width:80px" value=<%=rs("finalpaymentdate")%>>
		  		&nbsp;&nbsp;&nbsp;免箱期
		  		<input name="freestayperiod" style="width:100px" value=<%=rs("freestayperiod")%>>
		  		&nbsp;&nbsp;&nbsp;卫生证版本 
				<input name="weishengzhengversion" style="width:100px" value=<%=rs("weishengzhengversion")%>>
				&nbsp;&nbsp;&nbsp;放行日期
		  		<input name="passdate" id="passdate" class="datepicker" style="width:80px" value=<%=rs("passdate")%>>
				</td>
      </tr>       
      <tr>	  
	    	<td align="right" height="30">到单日期：</td>
        <td class="category">
		  		<input name="documentarrival" id="documentarrival" class="datepicker" style="width:80px" value=<%=rs("documentarrival")%>>
		  		&nbsp;&nbsp;&nbsp;预计赎单日期
		  		<input name="planretirebill" id="planretirebill" class="datepicker" style="width:80px" value=<%=rs("planretirebill")%>>	  		
		  		&nbsp;&nbsp;&nbsp;我司赎单日期
		  		<input name="internalretirebill" id="internalretirebill" class="datepicker" style="width:80px" value=<%=rs("internalretirebill")%>>		  		
		  		&nbsp;&nbsp;&nbsp;代理赎单日期
		  		<input name="externalretirebill" id="externalretirebill" class="datepicker" style="width:80px" value=<%=rs("externalretirebill")%>>		  		
				</td>
      </tr>       
      <tr>	  
	    	<td align="right" height="30">送货日期：</td>
        <td class="category">
		  		<input name="cargodate" id="cargodate" class="datepicker" style="width:80px" value=<%=rs("cargodate")%>>
		  		&nbsp;&nbsp;&nbsp;送货方向
		  		<input name="cargodir" style="width:150px" value=<%=rs("cargodir")%>>
		  		&nbsp;&nbsp;&nbsp;交单日期
		  		<input name="deliverydate" id="deliverydate" class="datepicker" style="width:80px" value=<%=rs("deliverydate")%>>	  		
				</td>
      </tr> 
      <tr>
        <td align="right" height="30" valign="top">单证问题：</td>
        <td class="category" valign="top">
        	<textarea name="documentissue" cols="30" rows="3" value=<%=rs("documentissue")%>></textarea>&nbsp;&nbsp;&nbsp;
        	<label for="documentcheck" style="vertical-align:top"> 审单情况 </label>
        	<textarea name="documentcheck" id="documentcheck" cols="30" rows="3" value=<%=rs("documentcheck")%>></textarea>&nbsp;&nbsp;&nbsp;
        	<label for="documentdelivery" style="vertical-align:top"> 寄单情况 </label>
        	<textarea name="documentdelivery" id="documentdelivery" cols="30" rows="3" value=<%=rs("documentdelivery")%>></textarea>&nbsp;&nbsp;&nbsp;
        	<label for="clearcustom" style="vertical-align:top">清关进程</label>
        	<textarea name="clearcustom" id="clearcustom" cols="30" rows="3" value=<%=rs("clearcustom")%>></textarea>            	   	
        </td>
      </tr>	 
      <tr>
        <td align="right" height="30" valign="top">保证金情况：</td>
        <td class="category" valign="top">
        	<textarea name="warrantyinformation" cols="30" rows="3" value=<%=rs("warrantyinformation")%>></textarea>&nbsp;&nbsp;&nbsp;
        	<label for="warrantystatus" style="vertical-align:top">保证金进度</label>
					<%
					sql="select * from warrantystatus"
					set rs_warrstatus=conn.execute(sql)
					%>
					<select name="warrantystatus"  style="vertical-align:top">
						<%
							do while rs_warrstatus.eof=false
						%>
							<option value="<%=rs_warrstatus("warrantystatus")%>" <%if rs("warrantystatus")=rs_warrstatus("warrantystatus") then%>selected="selected"<%end if%>><%=rs_warrstatus("warrantystatus")%></option>
						<%
							rs_warrstatus.movenext
							loop
						%>
					</select>	 
<!--					       	
					<select name="warrantystatus" style="vertical-align:top">
	 						<option value="审核中">审核中</option>	
	 						<option value="已退保">已退保</option>	
	 						<option value="已到账">已到账</option>	
					</select>			-->
					<label for="warranty" style="vertical-align:top">&nbsp;&nbsp;&nbsp;保证金金额</label>	
					<input name="warranty" style="vertical-align:top;width:100px" maxlength="15" value=<%=rs("warranty")%> onKeyPress="javascript:CheckNum();" onKeyUp="this.value=this.value.replace(/[^\d.]/g,'')">
					<label for="returnwarranty" style="vertical-align:top">&nbsp;&nbsp;&nbsp;退保证金金额</label>	
					<input name="returnwarranty" style="vertical-align:top;width:100px" maxlength="15" value=<%=rs("returnwarranty")%> onKeyPress="javascript:CheckNum();" onKeyUp="this.value=this.value.replace(/[^\d.]/g,'')">
        </td>
      </tr>	        
      <tr>
        <td align="right" height="30" valign="top">少货情况：</td>
        <td class="category" valign="top">
        	<textarea name="stockmiss" cols="30" rows="3" value=<%=rs("stockmiss")%>></textarea>&nbsp;&nbsp;&nbsp;
        	<label for="claiminformation" style="vertical-align:top">索赔情况</label>
        	<textarea name="claiminformation" id="claiminformation" cols="30" rows="3" value=<%=rs("claiminformation")%>></textarea>&nbsp;&nbsp;&nbsp;
					<label for="claimstatus" style="vertical-align:top">索赔状态</label>
					<%
					sql="select * from claimstatus"
					set rs_claimstatus=conn.execute(sql)
					%>
					<select name="claimstatus"  style="vertical-align:top">
						<%
							do while rs_claimstatus.eof=false
						%>
							<option value="<%=rs_claimstatus("claimstatus")%>" <%if rs("claimstatus")=rs_claimstatus("claimstatus") then%>selected="selected"<%end if%>><%=rs_claimstatus("claimstatus")%></option>
						<%
							rs_claimstatus.movenext
							loop
						%>
					</select>					
<!--
					<select name="claimstatus" style="vertical-align:top">
							<option value="索赔中">索赔中</option>
							<option value="需索赔">需索赔</option>
							<option value="不需索赔">不需索赔</option>
					</select>     -->
					<label for="claimprocessor" style="vertical-align:top">索赔负责人</label>
					<select name="claimprocessor" style="vertical-align:top">
	    			<% for i = 0 to customCount-1 %>
	 						<option value=<%=custom(i)%> <%if rs("claimprocessor")=custom(i) then%>selected="selected"<%end if%>><%=custom(i)%></option>
	 					<% next %>  			
					</select>&nbsp;&nbsp;&nbsp;								
        </td>
      </tr>	       
      <tr>
		  <input type="submit" value=" 确认修改 " onClick="return check1();" class="button">&nbsp;&nbsp;&nbsp;
		  <input type="hidden" name="hid1" value="ok">
		  <input type="button" value=" 放弃修改返回 " onClick="if (confirm('确定要放弃修改吗？')) {window.open('../master/delete_lock_table.asp?tablename=shipment&combinedkey=<%=request("shipment")%>'); window.location.href='shipment.asp';}" class="button">
      </tr>  
      </tbody>  
</table>  
</form>
<table align="center" cellpadding="4" cellspacing="1" class="toptable grid" border="1">
	  <form name="formitem">
      <tr>
		<input type="button" onclick="addNewRow()" value="增加一行" class="button"/>&nbsp;&nbsp;&nbsp;
		<input type="button" onclick="removeRow()" value="删除一行" class="button"/>
      </tr>	      
      <tr>
		<table align="center" cellpadding="4" cellspacing="1" class="toptable grid" border="1" id="item">
		  <tr>
		  	<td width="20"></td>			
	        <td width="30" align="center">项目号</td>
	        <td width="100" align="center">品名</td>
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
     	  <%
     	  	sql = "select * from shipmentitem where shipmentnum="&request("shipment")
     	  	'set rs_items=conn.execute(sql)
     	  	set rs_items =server.createobject("ADODB.RecordSet")	
 			 rs_items.open sql,conn,1,2
		    if not rs_items.eof then
		    	do while rs_items.eof=false     	  	
     	  %>	
		  <tr>
		  	<td width="20"><input type='checkbox' name='chkArr'  id='chkArr' /></td>			
	        <td width="30" align="center" ><input type='text' name='itemno' id='itemno' value='<%=rs_items("itemnum")%>' style='width:30px' readonly/></td>
	        <td width="100" align="center" >
	        	<select name="material<%=rs_items("itemnum")%>" id="material<%=rs_items("itemnum")%>" style='width:100px'>
	    			<% for i = 0 to materialCount-1 %>
	 					<option value='<%=material(i)%>' <%if rs_items("material")=material(i) then%>selected="selected"<%end if%>><%=material(i)%></option>
	 				<% next %>
	 			</select>
	        </td>
	        <td width="100" align="center" >
	        	<select name="customer<%=rs_items("itemnum")%>" id="customer<%=rs_items("itemnum")%>" style='width:100px'>
	    			<% for i = 0 to customerCount-1 %>
	 					<option value='<%=customer(i)%>' <%if rs_items("customer")=customer(i) then%>selected="selected"<%end if%>><%=customer(i)%></option>
	 				<% next %>
	 			</select>	        	
	        </td> 	
	        <td width=100" align="center" ><input type='text' name="spec<%=rs_items("itemnum")%>" id="spec<%=rs_items("itemnum")%>" value='<%=rs_items("spec")%>' style='width:100px'/></td> 		
	        <td width=80" align="center" ><input type='text' name="contractweight<%=rs_items("itemnum")%>" id="contractweight<%=rs_items("itemnum")%>" value='<%=rs_items("contractweight")%>' style='width:80px'/></td> 
	        <td width="40" align="center" >公斤</td> 		
	        <td width="80" align="center" ><input type='text' name="actualweight<%=rs_items("itemnum")%>" id="actualweight<%=rs_items("itemnum")%>" value='<%=rs_items("actualnetweight")%>' style='width:80px'/></td>
	        <td width="80" align="center" ><input type='text' name="purchaseprice<%=rs_items("itemnum")%>" id="purchaseprice<%=rs_items("itemnum")%>" value='<%=rs_items("purchaseprice")%>' style='width:80px'/></td> 	
	        <td width="40" align="center" >公斤</td> 				        	            		        	        		        
	        <td width="50" align="center" ><input name="producedate<%=rs_items("itemnum")%>" id="producedate<%=rs_items("itemnum")%>" value="<%=rs_items("productiondate")%>" style='width:80px' class="datepicker"/></td> 		
	        <td width="50" align="center" ><input type='text' name="casenum<%=rs_items("itemnum")%>" id="casenum<%=rs_items("itemnum")%>" value='<%=rs_items("casenumber")%>' style='width:80px'/></td> 		
	        <td width="100" align="center" ><input type='text' name="invamount<%=rs_items("itemnum")%>" id="invamount<%=rs_items("itemnum")%>" value='<%=rs_items("invoiceamount")%>' style='width:80px'/></td> 	
	        <td width="50" align="center" >
					<%
					sql="select * from trancurrency"
					set rs_currency=conn.execute(sql)
					%>
					<select name="invcurr<%=rs_items("itemnum")%>" style='width:50px'>
						<%
							do while rs_currency.eof=false
						%>
							<option value="<%=rs_currency("currency")%>" <%if rs_items("invoicecurrency")=rs_currency("currency") then%>selected="selected"<%end if%>><%=rs_currency("currency")%></option>
						<%
							rs_currency.movenext
							loop
						%>
					</select>	        	
	        </td> 		
	        <td width="100" align="center" ><input type='text' name="finalamount<%=rs_items("itemnum")%>" id="finalamount<%=rs_items("itemnum")%>" value='<%=rs_items("finalpayment")%>' style='width:80px'/></td>
     	  </tr>	
		  <%
		    	rs_items.movenext
		  		loop
		  	end if
		  %>     	  
        </table>			  	
      </tr>	          

  </form>
</table>
</form>
</td>
<td></td>
</tr> 
</table>

<%end if%>
</body>
</html>