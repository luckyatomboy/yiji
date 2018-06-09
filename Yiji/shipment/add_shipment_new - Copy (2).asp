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
<script type="text/javascript" src="../js/jquery-3.2.1.js"></script>
<script type="text/javascript" src="../js/jquery-ui.js"></script>
<!--<script type="text/javascript" src="../js/tableProcess.js"></script>-->
<!--<script type="text/javascript" src="../js/jquery.dataTables.min.js"></script>
<script type="text/javascript" src="../js/dataTables.editor.min.js"></script>-->

<link href="../style/jquery-ui.css" rel="stylesheet" type="text/css">
<link href="../style/style.css" rel="stylesheet" type="text/css">
<!--<link href="../style/jquery.dataTables.min.css" rel="stylesheet" type="text/css">
<link href="../style/editor.dataTables.min.css" rel="stylesheet" type="text/css">-->

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


<script>
//var editor; // use a global for the submit and return data rendering in the examples

	$(function(){
//日期控件。要支持动态创建的控件
		$("body").delegate(".datepicker","focusin",function(){
			$(this).datepicker();
		});
	});

//离开页面时提示
  $(window).bind('beforeunload',function(){
    return '数据尚未保存，确定离开?';
  });

</script>

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
'第一行
nowtrantype=request("trantype")
nowstatus=request("status")
nowinvoicestatus=request("invoicestatus")
'第二行
nowbuyer=request("buyer")
nowhandler=request("custom")
nowsales=request("sales")
'第三行
nowvendor=request("vendor")
nowplant=request("plant")
nowcountry=request("country")
nowincoterm=request("incoterm")
'第四行
nowcontract=request("contract")
nowmaterialtype=request("materialtype")
nowagent=request("agent")
nowpiwen=request("piwen")
if request("twodocumentready") <> "" then
  nowtwodocready="#"&request("twodocumentready")&"#"
else
  nowtwodocready="null"
end if
'第五行
nowport=request("port")
nowterminal=request("terminal")
nowcarrier=request("carrier")
nowshipname=request("shipname")
'第六行
nowdongjian=request("dongjian")
nowzidong=request("zidong")
if request("zidongreportdate") <> "" then
  nowzidongreportdate="#"&request("zidongreportdate")&"#"
else
  nowzidongreportdate="null"
end if
if request("zidongapplydate") <> "" then
  nowzidongapplydate="#"&request("zidongapplydate")&"#"
else
  nowzidongapplydate="null"
end if
'nowzidongreportdate=request("zidongreportdate")
'nowzidongapplydate=request("zidongapplydate")
'第七行
nowplaninsurance=request("planinsurance")
nowsupinsurance=request("supinsurance")
nowinsurancepayment=request("insurancepayment")
nowinsurancenumber=request("insurancenumber")
'第八行
nowplanship=request("planship")
nowshipdate=request("shipdate")
nowboarddate=request("boarddate")
nowcustomerdeliverydate=request("customerdeliverydate")
'第九行
'nowitemno=request("itemno").count
'nowcase=nowitemno
nowcase=request("case")
nowladnumber=request("ladnumber")
nowweishengzheng=request("weishengzheng")
nowlocknumber=request("locknumber")
if request("einformation")="有" then
  noweinformation="True"
else
  noweinformation="False"
end if
if request("MuslimCertification")="有" then
  nowMuslimCertification="True"
else
  nowMuslimCertification="False"
end if
if request("elabel")="有" then
  nowelabel="True"
else
  nowelabel="False"
end if
'第十行
nowprepaydate=request("prepaydate")
nowpayment=request("payment")
nowcurrency=request("currency")
nowdownpaymentdate=request("downpaymentdate")
nowtradingterm=request("tradingterm")
'第十一行
nowfinalpaymentdate=request("finalpaymentdate")
nowfreestayperiod=request("freestayperiod")
nowshouyi=request("shouyi")
nowpassdate=request("passdate")
'第十二行
nowdocumentarrival=request("documentarrival")
nowplanretirebill=request("planretirebill")
nowinternalretirebill=request("internalretirebill")
nowexternalretirebill=request("externalretirebill")
'第十三行
nowcargodate=request("cargodate")
nowcargodirection=request("cargodirection")
nowdeliverydate=request("deliverydate")
'第十四行
nowdocumentissue=request("documentissue")
nowdocumentcheck=request("documentcheck")
nowdocumentdelivery=request("documentdelivery")
nowclearcustom=request("clearcustom")
'第十五行
nowwarrantyinformation=request("warrantyinformation")
nowwarrantystatus=request("warrantystatus")
nowwarranty=request("warranty")
nowreturnwarranty=request("returnwarranty")
'第十六行
nowstockmiss=request("stockmiss")
nowclaiminformation=request("claiminformation")
nowclaimstatus=request("claimstatus")
nowclaimprocessor=request("claimprocessor")

temp_year = Year(now())
temp_month = Month(now())
if temp_month < 10 then
  temp_month = "0"&temp_month
end if
sql="select top 1 * from shipment where shipmentnum like '"&temp_year&temp_month&"%' order by shipmentnum desc"
set rs_count=conn.execute(sql)
if rs_count.bof and rs_count.eof then
  nowshipment=temp_year&temp_month&"001"
else
  rs_count.movefirst
  nowshipment=rs_count("shipmentnum") + 1
end if  

sql="insert into shipment(shipmentnum,trantype,status,invoicestatus,buyer,handler,sales,vendor,plant,country,incoterm,contract,materialtype,agent,piwen,twodocumentready,destination,terminal,carrier,shipname, dongjian,zidong,zidongapplydate,zidongreportdate,planinsurance,supinsurance, insurancepayment, insurancenumber, planship, shipdate, boarddate, customerdeliverydate, case, ladnumber, weishengzheng, locknumber, einformation, MuslimCertification, elabel, prepaydate, prepayment, trancurrency, downpaymentreceiptdate, tradingterm, finalpaymentdate, freestayperiod, shouyi, passdate, documentarrival, planretirebill, internalretirebill, externalretirebill, cargodate, cargodirection, deliverydate, documentissue, documentcheck, documentdelivery, clearcustom, warrantyinformation, warrantystatus, warranty, returnwarranty, stockmiss, claiminformation, claimstatus, claimprocessor, createdate,creator) values("&nowshipment&",'"&nowtrantype&"','"&nowstatus&"','"&nowinvoicestatus&"','"&nowbuyer&"','"&nowhandler&"','"&nowsales&"','"&nowvendor&"','"&nowplant&"','"&nowcountry&"','"&nowincoterm&"','"&nowcontract&"','"&nowmaterialtype&"','"&nowagent&"','"&nowpiwen&"',#"&nowtwodocready&"#,'"&nowportn&"','"&nowterminal&"','"&nowcarrier&"','"&nowshipname&"','"&nowdongjian&"','"&nowzidong&"',#"&nowzidongapplydate&"#,#"&nowzidongreportdate&"#,#"&nowplaninsurance&"#,#"&nowsupinsurance&"#,#"&nowinsurancepayment&"#,'"&nowinsurancenumber&"','"&nowplanship&"',#"&nowshipdate&"#,#"&nowboarddate&"#,#"&nowcustomerdeliverydate&"#,'"&nowcase&"','"&nowladnumber&"','"&nowweishengzheng&"','"&nowlocknumber&"',"&noweinformation&","&nowMuslimCertification&","&nowelabel&",#"&nowprepaydate&"#,"&nowpayment&",'"&nowcurrency&"',#"&nowdownpaymentdate&"#,'"&nowtradingterm&"',#"&nowfinalpaymentdate&"#,'"&nowfreestayperiod&"','"&nowshouyi&"',#"&nowpassdate&"#,#"&nowdocumentarrival&"#,#"&nowplanretirebill&"#,#"&nowinternalretirebill&"#,#"&nowexternalretirebill&"#,#"&nowcargodate&"#,'"&nowcargodirection&"',#"&nowdeliverydate&"#,'"&nowdocumentissue&"','"&nowdocumentcheck&"','"&nowdocumentdelivery&"','"&nowclearcustom&"','"&nowwarrantyinformation&"','"&nowwarrantystatus&"',"&nowwarranty&","&nowreturnwarranty&",'"&nowstockmiss&"','"&nowclaiminformation&"','"&nowclaimstatus&"','"&nowclaimprocessor&"',#"&now()&"#,'"&session("redboy_username")&"')"

'sql="insert into shipment(shipmentnum,trantype,status,invoicestatus,buyer,handler,sales,vendor,plant,country,incoterm,contract,materialtype,agent,piwen,twodocumentready,destination,terminal,carrier,shipname, dongjian,zidong,zidongapplydate,zidongreportdate,"

'if request("twodocumentready") <> "" then
  'sql=sql&"twodocumentready,"
'end if

'sql=sql&"destination,terminal,carrier,shipname, dongjian,zidong,"

'if request("zidongapplydate") <> "" thendestination,terminal,carrier,shipname, dongjian,zidong,
  'sql=sql&"zidongapplydate,"
'end if
'if request("zidongreportdate") <> "" then
  'sql=sql&"zidongreportdate,"
'end if 
if request("planinsurance") <> "" then
  sql=sql&"planinsurance,"
end if 
if request("supinsurance") <> "" then
  sql=sql&"supinsurance,"
end if 
if request("insurancepayment") <> "" then
  sql=sql&"insurancepayment,"
end if 

sql=sql&"insurancenumber, planship, "

if request("shipdate") <> "" then
  sql=sql&"shipdate,"
end if 
if request("boarddate") <> "" then
  sql=sql&"boarddate,"
end if 
if request("customerdeliverydate") <> "" then
  sql=sql&"customerdeliverydate,"
end if 

sql=sql&"case, ladnumber, weishengzheng, locknumber, einformation, MuslimCertification, elabel, "

if request("prepaydate") <> "" then
  sql=sql&"prepaydate,"
end if 

sql=sql&"prepayment, trancurrency, "

if request("downpaymentdate") <> "" then
  sql=sql&"downpaymentreceiptdate,"
end if 

sql=sql&"tradingterm, "

if request("finalpaymentdate") <> "" then
  sql=sql&"finalpaymentdate,"
end if 

sql=sql&"freestayperiod, shouyi, "

if request("passdate") <> "" then
  sql=sql&"passdate,"
end if 
if request("documentarrival") <> "" then
  sql=sql&"documentarrival,"
end if 
if request("planretirebill") <> "" then
  sql=sql&"planretirebill,"
end if 
if request("internalretirebill") <> "" then
  sql=sql&"internalretirebill,"
end if 
if request("externalretirebill") <> "" then
  sql=sql&"externalretirebill,"
end if 
if request("cargodate") <> "" then
  sql=sql&"cargodate,"
end if 
if request("deliverydate") <> "" then
  sql=sql&"deliverydate,"
end if 
sql=sql&"cargodirection, documentissue, documentcheck, documentdelivery, clearcustom, warrantyinformation, warrantystatus, warranty, returnwarranty, stockmiss, claiminformation, claimstatus, claimprocessor, createdate,creator) values("&nowshipment&",'"&nowtrantype&"','"&nowstatus&"','"&nowinvoicestatus&"','"&nowbuyer&"','"&nowhandler&"','"&nowsales&"','"&nowvendor&"','"&nowplant&"','"&nowcountry&"','"&nowincoterm&"','"&nowcontract&"','"&nowmaterialtype&"','"&nowagent&"','"&nowpiwen&"',"&nowtwodocready&",'"&nowport&"','"&nowterminal&"','"&nowcarrier&"','"&nowshipname&"','"&nowdongjian&"','"&nowzidong&"',"&nowzidongapplydate&","&nowzidongreportdate&","

'if request("twodocumentready") <> "" then
  'sql=sql&"#"&nowtwodocready&"#,"
'end if
'sql=sql&"'"&nowportn&"','"&nowterminal&"','"&nowcarrier&"','"&nowshipname&"','"&nowdongjian&"','"&nowzidong&"',"
'if request("zidongapplydate") <> "" then
  'sql=sql&"#"&nowzidongapplydate&"#,"
'end if
'if request("zidongreportdate") <> "" then
  'sql=sql&"#"&nowzidongreportdate&"#,"
'end if 
if request("planinsurance") <> "" then
  sql=sql&"#"&nowplaninsurance&"#,"
end if 
if request("supinsurance") <> "" then
  sql=sql&"#"&nowsupinsurance&"#,"
end if 
if request("insurancepayment") <> "" then
  sql=sql&"#"&nowinsurancepayment&"#,"
end if 
sql=sql&"'"&nowinsurancenumber&"','"&nowplanship&"',"
if request("shipdate") <> "" then
  sql=sql&"#"&nowshipdate&"#,"
end if 
if request("boarddate") <> "" then
  sql=sql&"#"&nowboarddate&"#,"
end if 
if request("customerdeliverydate") <> "" then
  sql=sql&"#"&nowcustomerdeliverydate&"#,"
end if 
sql=sql&"'"&nowcase&"','"&nowladnumber&"','"&nowweishengzheng&"','"&nowlocknumber&"',"&noweinformation&","&nowMuslimCertification&","&nowelabel&","
if request("prepaydate") <> "" then
  sql=sql&"#"&nowprepaydate&"#,"
end if 
sql=sql&nowpayment&",'"&nowcurrency&"',"
if request("downpaymentdate") <> "" then
  sql=sql&"#"&nowdownpaymentdate&"#,"
end if 
sql=sql&"'"&nowtradingterm&"',"
if request("finalpaymentdate") <> "" then
  sql=sql&"#"&nowfinalpaymentdate&"#,"
end if 
sql=sql&"'"&nowfreestayperiod&"','"&nowshouyi&"',"
if request("passdate") <> "" then
  sql=sql&"#"&nowpassdate&"#,"
end if 
if request("documentarrival") <> "" then
  sql=sql&"#"&nowdocumentarrival&"#,"
end if 
if request("planretirebill") <> "" then
  sql=sql&"#"&nowplanretirebill&"#,"
end if 
if request("internalretirebill") <> "" then
  sql=sql&"#"&nowinternalretirebill&"#,"
end if 
if request("externalretirebill") <> "" then
  sql=sql&"#"&nowexternalretirebill&"#,"
end if 
if request("cargodate") <> "" then
  sql=sql&"#"&nowcargodate&"#,"
end if 
if request("deliverydate") <> "" then
  sql=sql&"#"&nowdeliverydate&"#,"
end if 
sql=sql&"'"&nowcargodirection&"','"&nowdocumentissue&"','"&nowdocumentcheck&"','"&nowdocumentdelivery&"','"&nowclearcustom&"','"&nowwarrantyinformation&"','"&nowwarrantystatus&"',"&nowwarranty&","&nowreturnwarranty&",'"&nowstockmiss&"','"&nowclaiminformation&"','"&nowclaimstatus&"','"&nowclaimprocessor&"',#"&now()&"#,'"&session("redboy_username")&"')"

conn.execute(sql)

for i = 1 to request("itemno").count
nowitemno=i
nowmatitem=request("material"&i)
nowcusitem=request("customer"&i)
nowspecitem=request("spec"&i)
nowpackageitem=request("package"&i)
nowcontweightitem=request("contractweight"&i)
nowactualweightitem=request("actualweight"&i)
nowpurpriceitem=request("purchaseprice"&i)
nowproducedateitem=request("produceDate"&i)
nowcasenumitem=request("casenum"&i)
nowinvamtitem=request("invamount"&i)
nowcurritem=request("invcurr"&i)
nowfinalamtitem=request("finalamount"&i)
nowcontractweightuomitem=request("contractweightuom"&i)
nowpriceunititem=request("priceunit"&i)

'sql="insert into shipmentitem(shipmentnum,itemnum,material,customer,spec,contractweight,contractweightuom,purchaseprice,purchasepriceunit,productiondate,casenumber,actualnetweight,actualnetweightuom,invoiceamount,invoicecurrency,finalpayment,finalpaymentcurrency) values("&nowshipment&",'"&nowitemno&"','"&nowmatitem&"','"&nowcusitem&"','"&nowspecitem&"','"&nowcontweightitem&"','鹿芦陆茂','"&nowpurpriceitem&"','鹿芦陆茂','"&nowproducedateitem&"',"&nowcasenumitem&",'"&nowactualweightitem&"','鹿芦陆茂','"&nowinvamtitem&"','"&nowcurritem&"','"&nowfinalamtitem&"','"&nowcurritem&"')"
sql="insert into shipmentitem(shipmentnum,itemnum,material,customer,spec,package,contractweight,contractweightuom,actualnetweight,actualnetweightuom,purchaseprice,purchasepriceunit,"
if nowproducedateitem <> "" then
  sql=sql&"productiondate,"
end if 
sql=sql&"casenumber,invoiceamount,invoicecurrency,finalpayment,finalpaymentcurrency) values("&nowshipment&","&nowitemno&",'"&nowmatitem&"','"&nowcusitem&"','"&nowspecitem&"','"&nowpackageitem&"',"&nowcontweightitem&",'"&nowcontractweightuomitem&"',"&nowactualweightitem&",'"&nowcontractweightuomitem&"',"&nowpurpriceitem&",'"&nowpriceunititem&"',"
if nowproducedateitem <> "" then
  sql=sql&"#"&nowproducedateitem&"#,"
end if 
sql=sql&nowcasenumitem&","&nowinvamtitem&",'"&nowcurritem&"',"&nowfinalamtitem&",'"&nowcurritem&"')"
conn.execute(sql)

next

%>
<script language="javascript">
//alert(&nowshipment&)
alert("船期表录入成功！")
$(window).off('beforeunload');
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
   newTdObj1.innerHTML="<input type='checkbox' name='chkArr'  id='chkArr' align='middle'/>";
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
   newTdObj6.innerHTML="<input type='text' name='package"+itemNo+"' id='package"+itemNo+"' style='width:100px'/>";      
   var newTdObj7=myNewRow.insertCell(6);
   newTdObj7.innerHTML="<input type='text' name='contractWeight"+itemNo+"' id='contractWeight"+itemNo+"' style='width:80px' value=0 onKeyPress=\""+"javascript:CheckNum();\""+" onKeyUp=\""+"this.value=this.value.replace(/[^\\d.]/g,'')\""+">";        
   var newTdObj8=myNewRow.insertCell(7);
   newTdObj8.innerHTML="<select name='contractWeightUom"+itemNo+"' id='contractWeightUom"+itemNo+"' style='width:50px'>"
      +" <option value='公斤'>公斤</option>"
      +" <option value='磅'>磅</option>"
      +" </select>";       
   var newTdObj9=myNewRow.insertCell(8);
   newTdObj9.innerHTML="<input type='text' name='actualWeight"+itemNo+"' id='actualWeight"+itemNo+"' style='width:80px' value=0 onKeyPress=\""+"javascript:CheckNum();\""+" onKeyUp=\""+"this.value=this.value.replace(/[^\\d.]/g,'')\""+">";        
   var newTdObj10=myNewRow.insertCell(9);
   newTdObj10.innerHTML="<input type='text' name='purchasePrice"+itemNo+"' id='purchasePrice"+itemNo+"' style='width:80px' value=0 onKeyPress=\""+"javascript:CheckNum();\""+" onKeyUp=\""+"this.value=this.value.replace(/[^\\d.]/g,'')\""+">";       
   var newTdObj11=myNewRow.insertCell(10);
   newTdObj11.innerHTML="<select name='priceunit"+itemNo+"' id='priceunit"+itemNo+"' style='width:50px'>"
      +" <option value='公斤'>公斤</option>"
      +" <option value='磅'>磅</option>"
      +" </select>";      
   var newTdObj12=myNewRow.insertCell(11);
   newTdObj12.innerHTML="<input name='produceDate"+itemNo+"' id='produceDate"+itemNo+"' style='width:80px' class='datepicker'>";
//    + " <img src='../images/date.gif' align='absmiddle' style='cursor:pointer;'" 
//    + " onClick=\""+"JavaScript:window.open('../day.asp?form=formitem&field=produceDate"+itemNo+"&oldDate=produceDate"+itemNo+".value','','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');\"/>";     
   var newTdObj13=myNewRow.insertCell(12);    //箱数
   newTdObj13.innerHTML="<input name='casenum"+itemNo+"' id='casenum"+itemNo+"' style='width:80px' value=0 onKeyPress=\""+"javascript:CheckNum();\""+" onKeyUp=\""+"this.value=this.value.replace(/[^\\d.]/g,'')\""+">";    
   var newTdObj14=myNewRow.insertCell(13);    //发票总金额
   newTdObj14.innerHTML="<input name='invAmount"+itemNo+"' id='invAmt"+itemNo+"' style='width:80px' value=0 onKeyPress=\""+"javascript:CheckNum();\""+" onKeyUp=\""+"this.value=this.value.replace(/[^\\d.]/g,'')\""+">";        
   var newTdObj15=myNewRow.insertCell(14);
   newTdObj15.innerHTML="<select name='invCurr"+itemNo+"' id='invCurr"+itemNo+"' style='width:50px'>"
      +" <option value='美元'>美元</option>"
      +" <option value='澳元'>澳元</option>"
      +" </select>";   
   var newTdObj16=myNewRow.insertCell(15);    //尾款金额
   newTdObj16.innerHTML="<input name='finalAmount"+itemNo+"' id='finalAmt"+itemNo+"' style='width:80px' value=0 onKeyPress=\""+"javascript:CheckNum();\""+" onKeyUp=\""+"this.value=this.value.replace(/[^\\d.]/g,'')\""+">";        

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
      <td>&nbsp;船期表</td>
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
	    	<td align="right" height="30">采购</td>
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
	    	<td align="right" height="30">供应商</td>
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
					&nbsp;&nbsp;&nbsp;&nbsp;国家			
					<select name="country">			
						<option selected="selected">---请选择---</option>		  			
					</select>						
					&nbsp;&nbsp;&nbsp;&nbsp;工厂					
					<select name="plant"> <!-- onchange="this.selectedIndex=this.defaultIndex;">-->
						  <option selected="selected">---请选择---</option>		
					</select>					
		  		&nbsp;&nbsp;&nbsp;付款条件
<!--		  		<input name="incoterm" style="width:150px">							 -->
					<select name="incoterm">			
						  <option selected="selected">---脟毛脩隆脭帽---</option>										
					</select>				  				
				</td>
      </tr>
      <tr>	  
	    	<td align="right" height="30">合同号</td>
        <td class="category">
		  		<input name="contract" value="<%=contract%>" style="width:150px" maxlength="20">
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
							<option value=<%=rs_mattype("materialtype")%>><%=rs_mattype("materialtype")%></option>
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
					&nbsp;&nbsp;&nbsp;两证齐备日
		  		<input name="twodocumentready" id="twodocumentready" class="datepicker" style="width:80px">									 		
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
							<option value=<%=rs_port("port")%>><%=rs_port("port")%></option>
						<%
							rs_port.movenext
							loop
							rs_port.close
						%>											
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
							<option value=<%=rs_terminal("terminal")%>><%=rs_terminal("terminal")%></option>
						<%
							rs_terminal.movenext
							loop
							rs_terminal.close
						%>	
<!--						  <option value="脤矛陆貌脥芒麓煤">脤矛陆貌脥芒麓煤</option>		
						  <option value="脥芒脦氓">脥芒脦氓</option>		
						  <option value="脩贸脪禄">脩贸脪禄</option>		
						  <option value="脩贸脠媒">脩贸脠媒</option>				-->						
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
							<option value=<%=rs_carrier("carrier")%>><%=rs_carrier("carrier")%></option>
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
					<input name="shipname" style="width:150px" maxlength="20">							
				</td>
      </tr>               
      <tr>	  
	    	<td align="right" height="30">动检证：</td>
        <td class="category">
          <input name="dongjian" style="width:150px" maxlength="20">    
          &nbsp;&nbsp;&nbsp;&nbsp;自动证
          <input name="zidong" style="width:150px" maxlength="20">    
		  		&nbsp;&nbsp;&nbsp;&nbsp;自动证交批日期
		  		<input name="zidongapplydate" id="zidongapplydate" class="datepicker" style="width:80px">		  				  		
		  		&nbsp;&nbsp;&nbsp;&nbsp;自动证上报日期
		  		<input name="zidongreportdate" id="zidongreportdate" class="datepicker" style="width:80px">	  				  		
				</td>
      </tr>
      <tr>	  
	    	<td align="right" height="30">预保日：</td>
        <td class="category">
		  		<input name="planinsurance" id="planinsurance" class="datepicker" style="width:80px">
		  		&nbsp;&nbsp;&nbsp;&nbsp;补保日
		  		<input name="supinsurance" id="supinsurance" class="datepicker" style="width:80px">
		  		&nbsp;&nbsp;&nbsp;&nbsp;保费支付日
		  		<input name="insurancepayment" id="insurancepayment" class="datepicker" style="width:80px">	  		
		  		&nbsp;&nbsp;&nbsp;&nbsp;保单
		  		<input name="insurancenumber" style="width:120px">
				</td>
      </tr>      
      <tr>	  
	    	<td align="right" height="30">预计装船月份：</td>
        <td class="category">
		  		<input name="planship" style="width:100px">
		  		&nbsp;&nbsp;&nbsp;&nbsp;实际装船期
		  		<input name="shipdate" id="shipdate" class="datepicker" style="width:80px">
		  		&nbsp;&nbsp;&nbsp;&nbsp;预计到港期
		  		<input name="boarddate" id="boarddate" class="datepicker" style="width:80px">	  		
		  		&nbsp;&nbsp;&nbsp;&nbsp;客户交货期
		  		<input name="customerdeliverydate" class="datepicker" style="width:100px">
				</td>
      </tr>
      <tr>	  
	    	<td align="right" height="30">箱号：</td>
        <td class="category">
		  		<input name="case" style="width:120px" maxlength="15">
		  		&nbsp;&nbsp;&nbsp;提单号
		  		<input name="ladnumber" style="width:120px" maxlength="20">
		  		&nbsp;&nbsp;&nbsp;卫生证号
		  		<input class="bubble" name="weishengzheng" style="width:120px" maxlength="20" onClick="add_memo(this)">	
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
          &nbsp;&nbsp;&nbsp;标签
          <select name="ELabel">     
              <option value="有">有</option>    
              <option value="无">无</option>                    
          </select>               				  		
				</td>
      </tr>
      <tr>	  
	    	<td align="right" height="30">预付款日期：</td>
        <td class="category">
		  		<input name="prepaydate" id="prepaydate" class="datepicker" style="width:80px">
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
				<input name="downpaymentdate" id="downpaymentdate" class="datepicker" style="width:80px">	
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
		  		<input name="finalpaymentdate" id="finalpaymentdate" class="datepicker" style="width:80px">
		  		&nbsp;&nbsp;&nbsp;免箱期
		  		<input name="freestayperiod" style="width:100px">
		  		&nbsp;&nbsp;&nbsp;兽医
					<input name="shouyi" style="width:100px">
					&nbsp;&nbsp;&nbsp;放行日期
		  		<input name="passdate" id="passdate" class="datepicker" style="width:80px">
				</td>
      </tr>     
 
      <tr>	  
	    	<td align="right" height="30">到单日期：</td>
        <td class="category">
		  		<input name="documentarrival" id="documentarrival" class="datepicker" style="width:80px">
		  		&nbsp;&nbsp;&nbsp;预计赎单日期
		  		<input name="planretirebill" id="planretirebill" class="datepicker" style="width:80px">		  		
		  		&nbsp;&nbsp;&nbsp;我司赎单日期
		  		<input name="internalretirebill" id="internalretirebill" class="datepicker" style="width:80px">	  		
		  		&nbsp;&nbsp;&nbsp;代理赎单日期
		  		<input name="externalretirebill" id="externalretirebill" class="datepicker" style="width:80px">	  		
				</td>
      </tr>       
      <tr>	  
	    	<td align="right" height="30">送货日期：</td>
        <td class="category">
		  		<input name="cargodate" id="cargodate" class="datepicker" style="width:80px">
		  		&nbsp;&nbsp;&nbsp;送货方向
		  		<input name="cargodirection" style="width:150px">
		  		&nbsp;&nbsp;&nbsp;交单日期
		  		<input name="deliverydate" id="deliverydate" class="datepicker" style="width:80px">	  		
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
					<select name="claimprocessor" style="vertical-align:top">
	    			<% for i = 0 to customCount-1 %>
	 						<option value=<%=custom(i)%>><%=custom(i)%></option>
	 					<% next %>  			
					</select>								
        </td>
      </tr>	       
      <tr>
		  <input type="submit" value=" 确认录入 " onClick="$(window).off('beforeunload'); return check1();" class="button">&nbsp;&nbsp;&nbsp;
		  <input type="hidden" name="hid1" value="ok">
		  <input type="reset" value=" 重新填写 " class="button">
      </tr>    

</table>  
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
              <td width="30" align="center" >项目号</td>
              <td width="100" align="center" >品名</td>
              <td width="100" align="center" >客户</td>   
              <td width=100" align="center" >规格</td>    
              <td width=100" align="center" >包装</td>    
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