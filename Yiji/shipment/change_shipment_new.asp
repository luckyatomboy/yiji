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
<title><%=dianming%> - �޸Ĵ��ڱ�</title>
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
if fla1="0" and fla2="0" and fla3="0" and fla4="0" then
%>
<br><center><img src="../images/note.gif" align="absmiddle">&nbsp;<font color="#FF0000">�㲻�߱���Ȩ�ޣ��������Ա��ϵ��</font></center>
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
    alert("�û�<%=rs_lock("username")%>���ڱ༭�ü�¼�����Ժ����ԣ�");
    window.location.href="shipment.asp";
    </script> 
<%end if
else
    sql="insert into locktable(tablename,combinedkey,status,username,locktime) values('shipment','"&request("shipment")&"','E','"&session("redboy_username")&"',#"&now()&"#)"  
    conn.execute(sql)
end if
%>

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
if request("dongjian")="����ѡ�񶯼�֤" then
  rs("dongjian")=""
else
  rs("dongjian")=request("dongjian")
end if
rs("zidongcom")=request("zidongcom")
if request("zidong")="����ѡ���Զ�֤" then
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

'��ɾ�����е�item'
sql="delete from shipmentitem where shipmentnum="&request("shipment")
conn.execute(sql)

'�ٲ���item'
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

sql="insert into shipmentitem(shipmentnum,itemnum,material,customer,spec,contractweight,contractweightuom,actualnetweight,actualnetweightuom,purchaseprice,purchasepriceunit,productiondate,casenumber,invoiceamount,invoicecurrency,finalpayment,finalpaymentcurrency) values("&request("shipment")&",'"&nowitemno&"','"&nowmatitem&"','"&nowcusitem&"','"&nowspecitem&"','"&nowcontweightitem&"','����','"&nowactualweightitem&"','����','"&nowpurpriceitem&"','����',#"&nowproducedateitem&"#,'"&nowcasenumitem&"',"&nowinvamtitem&",'"&nowcurritem&"',"&nowfinalamtitem&",'"&nowcurritem&"')"

'sql="insert into shipmentitem(shipmentnum,itemnum,material,customer,spec,contractweight,contractweightuom,actualnetweight,actualnetweightuom) values("&request("shipment")&","&nowitemno&",'"&nowmatitem&"','"&nowcusitem&"','"&nowspecitem&"',"&nowcontweightitem&",'����',"&nowactualweightitem&",'����')"
conn.execute(sql)

next

'ɾ���߼���'
sql="delete from locktable where tablename='shipment' and combinedkey='"&request("shipment")&"'"
conn.execute(sql)

%>
<script language="javascript">
//alert(&nowshipment&)
alert("���ڱ�¼��ɹ���")
window.location.href="shipment.asp"
</script>

<%
else
%>
<script language="javascript">

<% 
'�������ݱ��浽���� 
dim vendorCount
sql="select * from vendor order by vendorname"
set rs_vendor=conn.execute(sql)
%> 
var country = new Array(); 
var plant = new Array();
var incoterm = new Array();
//����ṹ��һ����ֵ,������ֵ,������ʾֵ 
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
'����Ʒ�����ݵ�����
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
'����ͻ����ݵ�����
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
'�����û����ݵ�����
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
<!-- ����ͬ�� -->
if (document.form1.contract.value=="")
{
alert("�������ͬ�ţ�");
return false;
}
<!-- ����Զ�֤ -->
<!-- ����Զ�֤ -->
}


function chsel(vendor){
//���ù���	
    document.form1.country.length = 0; 
    for (i=0; i<country.length; i++) 
    { 
        if (country[i][0]==vendor) 
        {document.form1.country.options[0] = new Option(country[i][1]);} 
    } 
//���ù���    
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
//���ø�������   
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
 
  //���ڱ������һ��
  function addNewRow(){
   var tabObj=document.getElementById("item");//��ȡ������ݵı��
   var rowsNum = tabObj.rows.length;  //��ȡ��ǰ����
   var colsNum=tabObj.rows[0].cells.length;//��ȡ�е�����
   var myNewRow = tabObj.insertRow(rowsNum);//��������.
      
   var itemObj=document.getElementsByName("itemno");//ȡ�������е�itemno
   var itemNo=1;
   if (itemObj.length==0) {
  	 itemNo=1;//���û��item,��1
  	 
  }else{
  	 itemNo=parseInt(itemObj[itemObj.length-1].value) + 1; //ȡ�����itemno��1
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
   newTdObj7.innerHTML="����";   
   newTdObj7.align="center";
   var newTdObj8=myNewRow.insertCell(7);
   newTdObj8.innerHTML="<input type='text' name='actualWeight"+itemNo+"' id='actualWeight"+itemNo+"' style='width:80px'  align='center' value=0 onKeyPress=\""+"javascript:CheckNum();\""+" onKeyUp=\""+"this.value=this.value.replace(/[^\\d.]/g,'')\""+">";   
   newTdObj8.align="center";		 
   var newTdObj9=myNewRow.insertCell(8);
   newTdObj9.innerHTML="<input type='text' name='purchasePrice"+itemNo+"' id='purchasePrice"+itemNo+"' style='width:80px' align='center' value=0 onKeyPress=\""+"javascript:CheckNum();\""+" onKeyUp=\""+"this.value=this.value.replace(/[^\\d.]/g,'')\""+">";   	
   newTdObj9.align="center";	 
   var newTdObj10=myNewRow.insertCell(9);
   newTdObj10.innerHTML="����";     
   newTdObj10.align="center";
   var newTdObj11=myNewRow.insertCell(10);
   newTdObj11.innerHTML="<input name='produceDate"+itemNo+"' id='produceDate"+itemNo+"' readonly align='center' style='width:80px'"
//   	+ " <img src='../images/date.gif' align='absmiddle' style='cursor:pointer;'" 
   	+ " onClick=\""+"JavaScript:window.open('../day.asp?form=form1&field=produceDate"+itemNo+"&oldDate=produceDate"+itemNo+".value','','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');\"/>";   
   newTdObj11.align="center";  
   var newTdObj12=myNewRow.insertCell(11);   	//����
	 newTdObj12.innerHTML="<input name='casenum"+itemNo+"' id='casenum"+itemNo+"' style='width:80px' align='center' onKeyPress=\""+"javascript:CheckNum();\""+" onKeyUp=\""+"this.value=this.value.replace(/[^\\d.]/g,'')\""+">";   
   newTdObj12.align="center";	
   var newTdObj13=myNewRow.insertCell(12);   	//��Ʊ�ܽ��
	 newTdObj13.innerHTML="<input name='invAmount"+itemNo+"' id='invAmt"+itemNo+"' style='width:80px' align='center' value=0 onKeyPress=\""+"javascript:CheckNum();\""+" onKeyUp=\""+"this.value=this.value.replace(/[^\\d.]/g,'')\""+">";   	
   newTdObj13.align="center";	 
   var newTdObj14=myNewRow.insertCell(13);
   newTdObj14.innerHTML="<select name='invCurr"+itemNo+"' id='invCurr"+itemNo+"' style='width:50px' align='center'>"
	 		+" <option value='��Ԫ'>��Ԫ</option>"
	 		+" <option value='�ı�'>�ı�</option>"
	 		+" </select>";   
   newTdObj14.align="center";
   var newTdObj15=myNewRow.insertCell(14);   	//β����
	 newTdObj15.innerHTML="<input name='finalAmount"+itemNo+"' id='finalAmt"+itemNo+"' style='width:80px' value=0 align='center' onKeyPress=\""+"javascript:CheckNum();\""+" onKeyUp=\""+"this.value=this.value.replace(/[^\\d.]/g,'')\""+">";   
   newTdObj15.align="center";		 

  }

//���ڱ��ɾ��һ��
  function removeRow(){
   var chkObj=document.getElementsByName("chkArr");
   var tabObj=document.getElementById("item");
   var shipment=document.getElementById("shipment");
   var itemno=document.getElementsByName("itemno");
   var existing=0;

   for(var k=0;k<chkObj.length;k++){
    if(chkObj[k].checked){
//����Ƿ��Ѿ��н��㵥������¼������У�����ɾ��
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
    	alert("�Ѿ��н��㵥������ɾ������Ŀ��")
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
      <td width="10%" align="left">&nbsp;�޸Ĵ��ڱ�����(��)  ---  </td>
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
<table align="center" cellpadding="4" cellspacing="1" class="toptable grid" border="1">
<!--	  <form name="form1">	-->
      <tr>
        <td width="20%" align="right" height="30">���ڱ���룺</td>
        <td width="80%" class="category"><input name="shipment" id="shipment" value=<%=rs("shipmentnum")%> readonly></td>
      </tr>
      <tr>
        <td width="20%" align="right" height="30">һ����Ϣ��</td>
        <td width="80%" class="category">
        	���ͣ�
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
					&nbsp;&nbsp;&nbsp;&nbsp;״̬
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
					&nbsp;&nbsp;&nbsp;&nbsp;���㵥״̬
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
	    	<td align="right" height="30">�ɹ���</td>
        <td class="category">		
					<select name="buyer">
	    			<% for i = 0 to buyerCount-1 %>
	 						<option value="<%=buyer(i)%>" <%if buyer(i)=rs("buyer") then%>selected="selected"<%end if%>><%=buyer(i)%></option>
	 					<% next %>  			
					</select>						
					&nbsp;&nbsp;&nbsp;&nbsp;����	      
					<select name="custom">
	    			<% for i = 0 to customCount-1 %>
	 						<option value="<%=custom(i)%>" <%if custom(i)=rs("handler") then%>selected="selected"<%end if%>><%=custom(i)%></option>
	 					<% next %>  			
					</select>				
					&nbsp;&nbsp;&nbsp;&nbsp;����	      
					<select name="sales">
	    			<% for i = 0 to salesCount-1 %>
	 						<option value="<%=sales(i)%>" <%if sales(i)=rs("sales") then%>selected="selected"<%end if%>><%=sales(i)%></option>
	 					<% next %>  			
					</select>									  		
				</td>
      </tr>			
      <tr>	  
	    	<td align="right" height="30">��Ӧ�̣�</td>
        <td class="category">
					<%
					sql="select * from vendor order by vendorname"
					set rs_vendor=conn.execute(sql)
					%>
					<select name="vendor" onchange="chsel(this.value)">
						<option selected="selected">---��ѡ��---</option>	
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
					&nbsp;&nbsp;&nbsp;&nbsp;����				
					<select name="country">			
						<option selected="selected"><%=rs("country")%></option>		  			
					</select>						
					&nbsp;&nbsp;&nbsp;&nbsp;����						
					<select name="plant"> <!-- onchange="this.selectedIndex=this.defaultIndex;">-->
						  <option selected="selected"><%=rs("plant")%></option>		
					</select>					
		  		&nbsp;&nbsp;&nbsp;���ʽ
<!--		  		<input name="incoterm" style="width:150px">							 -->
					<select name="incoterm">			
						  <option selected="selected"><%=rs("incoterm")%></option>										
					</select>				  				
				</td>
      </tr>
      <tr>	  
	    	<td align="right" height="30">��ͬ�ţ�</td>
        <td class="category">
		  		<input name="contract" value=<%=rs("contract")%> style="width:150px" maxlength="20">
		  		&nbsp;<font color="#ff0000">*</font>
		  		&nbsp;&nbsp;&nbsp;&nbsp;Ʒ��
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
		  		&nbsp;&nbsp;&nbsp;&nbsp;����˾
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
		  		&nbsp;&nbsp;&nbsp;&nbsp;��������
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
					&nbsp;&nbsp;&nbsp;��֤
		  		<input name="twodocumentsready" id="twodocumentsready" value="<%=rs("twodocumentready")%>" readonly style="width:80px">
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=twodocumentsready&oldDate='+twodocumentsready.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">										 		
				</td>
      </tr>    
      <tr>	  
	    	<td align="right" height="30">�ۿڣ�</td>
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
<!--						  <option value="�Ϻ�">�Ϻ�</option>		
						  <option value="����">����</option>		
						  <option value="���">���</option>		
						  <option value="����">����</option>				-->						
					</select>		
					&nbsp;&nbsp;&nbsp;&nbsp;������ͷ
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
<!--						  <option value="������">������</option>		
						  <option value="����">����</option>		
						  <option value="��һ">��һ</option>		
						  <option value="����">����</option>				-->						
					</select>			
					&nbsp;&nbsp;&nbsp;&nbsp;����˾
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
					&nbsp;&nbsp;&nbsp;&nbsp;��������
					<input name="shipname" style="width:150px" maxlength="20">							
				</td>
      </tr>               
      <tr>	  
	    	<td align="right" height="30">����֤��</td>
        <td class="category">
		  		<input name="dongjian" readonly onClick="JavaScript:window.open('../huiyuan/query_license.asp?queryform=form1&licensetype=����֤&field=dongjian&field2=dongjiancom','','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=1200,height=470,top=100,left=20');" style="cursor:hand;width:120px" value=<%=rs("dongjian")%>>
		  		<input name="dongjiancom" readonly value=<%=rs("dongjiancom")%>>
				</td>
      </tr>
      <tr>	  
	    	<td align="right" height="30">�Զ�֤��</td>
        <td class="category">
		  		<input name="zidong" readonly onClick="JavaScript:window.open('../huiyuan/query_license.asp?queryform=form1&licensetype=�Զ�֤&field=zidong&field2=zidongcom','','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=1200,height=470,top=100,left=20');" style="cursor:hand;width:120px" value=<%=rs("zidong")%>>
		  		<input name="zidongcom" readonly value=<%=rs("zidongcom")%>>
		  		&nbsp;&nbsp;&nbsp;&nbsp;�Զ�֤��������
		  		<input name="zidongapplydate" id="zidongapplydate" readonly style="width:80px" value=<%=rs("zidongapplydate")%>>
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=zidongapplydate&oldDate='+zidongapplydate.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">		  				  		
		  		&nbsp;&nbsp;&nbsp;&nbsp;�Զ�֤�ϱ�����
		  		<input name="zidongreportdate" id="zidongreportdate" readonly style="width:80px" value=<%=rs("zidongreportdate")%>>
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=zidongreportdate&oldDate='+zidongreportdate.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">		  				  		
				</td>
      </tr>
      <tr>	  
	    	<td align="right" height="30">Ԥ���գ�</td>
        <td class="category">
		  		<input name="planinsurance" id="planinsurance" style="width:80px" value=<%=rs("planinsurance")%>>
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=planinsurance&oldDate='+planinsurance.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">
		  		&nbsp;&nbsp;&nbsp;&nbsp;������
		  		<input name="supinsurance" id="supinsurance" readonly style="width:80px" value=<%=rs("supinsurance")%>>
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=supinsurance&oldDate='+supinsurance.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">
		  		&nbsp;&nbsp;&nbsp;&nbsp;����֧����
		  		<input name="insurancepayment" id="insurancepayment" readonly style="width:80px" value=<%=rs("insurancepayment")%>>
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=insurancepayment&oldDate='+insurancepayment.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">		  		
		  		&nbsp;&nbsp;&nbsp;&nbsp;����
		  		<input name="insurancenumber" style="width:120px" value=<%=rs("insurancenumber")%>>
				</td>
      </tr>      
      <tr>	  
	    	<td align="right" height="30">Ԥ��װ���·ݣ�</td>
        <td class="category">
		  		<input name="planship" style="width:100px">
		  		&nbsp;&nbsp;&nbsp;&nbsp;ʵ��װ����
		  		<input name="shipdate" id="shipdate" readonly style="width:80px" value=<%=rs("shipdate")%>>
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=shipdate&oldDate='+shipdate.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">
		  		&nbsp;&nbsp;&nbsp;&nbsp;Ԥ�Ƶ�����
		  		<input name="boarddate" id="boarddate" readonly style="width:80px" value=<%=rs("boarddate")%>>
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=boarddate&oldDate='+boarddate.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">		  		
		  		&nbsp;&nbsp;&nbsp;&nbsp;�ͻ�������
		  		<input name="plandeliverydate" style="width:100px" value=<%=rs("plandeliverydate")%>>
				</td>
      </tr>
      <tr>	  
	    	<td align="right" height="30">��ţ�</td>
        <td class="category">
		  		<input name="case" style="width:120px" maxlength="15" value=<%=rs("case")%>>
		  		&nbsp;&nbsp;&nbsp;�ᵥ��
		  		<input name="ladnumber" style="width:120px" maxlength="20" value=<%=rs("ladnumber")%>>
		  		&nbsp;&nbsp;&nbsp;����֤��
		  		<input name="weishengzheng" style="width:120px" maxlength="20" value=<%=rs("weishengzheng")%>>	
					&nbsp;&nbsp;&nbsp;Ǧ���
					<input name="locknumber" style="width:120px" maxlength="20" value=<%=rs("locknumber")%>>		
		  		&nbsp;&nbsp;&nbsp;������Ϣ
					<select name="einformation">			
						  <option value="��" <%if rs("einformation")="True" then%>selected="selected"<%end if%>>��</option>		
						  <option value="��" <%if rs("einformation")="False" then%>selected="selected"<%end if%>>��</option>										
					</select>	
		  		&nbsp;&nbsp;&nbsp;����֤��
					<select name="MuslimCertification">			
						  <option value="��" <%if rs("MuslimCertification")="True" then%>selected="selected"<%end if%>>��</option>		
						  <option value="��" <%if rs("MuslimCertification")="False" then%>selected="selected"<%end if%>>��</option>										
					</select>									 		  		
				</td>
      </tr>
      <tr>	  
	    	<td align="right" height="30">Ԥ�������ڣ�</td>
        <td class="category">
		  		<input name="paydate" id="paydate" readonly style="width:80px" value=<%=rs("prepaydate")%>>
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=paydate&oldDate='+paydate.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">
		  		&nbsp;&nbsp;&nbsp;Ԥ������
		  		<input name="prepayment" style="width:100px" maxlength="15" value=<%=rs("prepayment")%> onKeyPress="javascript:CheckNum();" onKeyUp="this.value=this.value.replace(/[^\d.]/g,'')">
		  		&nbsp;&nbsp;&nbsp;����
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
				&nbsp;&nbsp;&nbsp;�����յ�����
				<input name="downpaymentdate" id="downpaymentdate" readonly style="width:80px" value=<%=rs("downpaymentreceiptdate")%>>	
				<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=downpaymentdate&oldDate='+downpaymentdate.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">
				&nbsp;&nbsp;&nbsp;�ɽ�����
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
	    	<td align="right" height="30">β��֧�����ڣ�</td>
        <td class="category">
		  		<input name="finalpaymentdate" id="finalpaymentdate" readonly style="width:80px" value=<%=rs("finalpaymentdate")%>>
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=finalpaymentdate&oldDate='+finalpaymentdate.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">
		  		&nbsp;&nbsp;&nbsp;������
		  		<input name="freestayperiod" style="width:100px" value=<%=rs("freestayperiod")%>>
		  		&nbsp;&nbsp;&nbsp;����֤�汾 
				<input name="weishengzhengversion" style="width:100px" value=<%=rs("weishengzhengversion")%>>
				&nbsp;&nbsp;&nbsp;��������
		  		<input name="passdate" id="passdate" readonly style="width:80px" value=<%=rs("passdate")%>>
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=passdate&oldDate='+passdate.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">
				</td>
      </tr>       
      <tr>	  
	    	<td align="right" height="30">�������ڣ�</td>
        <td class="category">
		  		<input name="documentarrival" id="documentarrival" readonly style="width:80px" value=<%=rs("documentarrival")%>>
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=documentarrival&oldDate='+documentarrival.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">
		  		&nbsp;&nbsp;&nbsp;Ԥ���굥����
		  		<input name="planretirebill" id="planretirebill" readonly style="width:80px" value=<%=rs("planretirebill")%>>
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=planretirebill&oldDate='+planretirebill.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">		  		
		  		&nbsp;&nbsp;&nbsp;��˾�굥����
		  		<input name="internalretirebill" id="internalretirebill" readonly style="width:80px" value=<%=rs("internalretirebill")%>>
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=internalretirebill&oldDate='+internalretirebill.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">		  		
		  		&nbsp;&nbsp;&nbsp;�����굥����
		  		<input name="externalretirebill" id="externalretirebill" readonly style="width:80px" value=<%=rs("externalretirebill")%>>
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=externalretirebill&oldDate='+externalretirebill.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">		  		
				</td>
      </tr>       
      <tr>	  
	    	<td align="right" height="30">�ͻ����ڣ�</td>
        <td class="category">
		  		<input name="cargodate" id="cargodate" readonly style="width:80px" value=<%=rs("cargodate")%>>
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=cargodate&oldDate='+cargodate.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">
		  		&nbsp;&nbsp;&nbsp;�ͻ�����
		  		<input name="cargodir" style="width:150px" value=<%=rs("cargodir")%>>
		  		&nbsp;&nbsp;&nbsp;��������
		  		<input name="deliverydate" id="deliverydate" readonly style="width:80px" value=<%=rs("deliverydate")%>>
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=deliverydate&oldDate='+deliverydate.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">		  		
				</td>
      </tr> 
      <tr>
        <td align="right" height="30" valign="top">��֤���⣺</td>
        <td class="category" valign="top">
        	<textarea name="documentissue" cols="30" rows="3" value=<%=rs("documentissue")%>></textarea>&nbsp;&nbsp;&nbsp;
        	<label for="documentcheck" style="vertical-align:top"> ����� </label>
        	<textarea name="documentcheck" id="documentcheck" cols="30" rows="3" value=<%=rs("documentcheck")%>></textarea>&nbsp;&nbsp;&nbsp;
        	<label for="documentdelivery" style="vertical-align:top"> �ĵ���� </label>
        	<textarea name="documentdelivery" id="documentdelivery" cols="30" rows="3" value=<%=rs("documentdelivery")%>></textarea>&nbsp;&nbsp;&nbsp;
        	<label for="clearcustom" style="vertical-align:top">��ؽ���</label>
        	<textarea name="clearcustom" id="clearcustom" cols="30" rows="3" value=<%=rs("clearcustom")%>></textarea>            	   	
        </td>
      </tr>	 
      <tr>
        <td align="right" height="30" valign="top">��֤�������</td>
        <td class="category" valign="top">
        	<textarea name="warrantyinformation" cols="30" rows="3" value=<%=rs("warrantyinformation")%>></textarea>&nbsp;&nbsp;&nbsp;
        	<label for="warrantystatus" style="vertical-align:top">��֤�����</label>
					<%
					sql="select * from warrantystatus"
					set rs_warrstatus=conn.execute(sql)
					%>
					<select name="warrantystatus">
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
	 						<option value="�����">�����</option>	
	 						<option value="���˱�">���˱�</option>	
	 						<option value="�ѵ���">�ѵ���</option>	
					</select>			-->
					<label for="warranty" style="vertical-align:top">&nbsp;&nbsp;&nbsp;��֤����</label>	
					<input name="warranty" style="vertical-align:top;width:100px" maxlength="15" value=<%=rs("warranty")%> onKeyPress="javascript:CheckNum();" onKeyUp="this.value=this.value.replace(/[^\d.]/g,'')">
					<label for="returnwarranty" style="vertical-align:top">&nbsp;&nbsp;&nbsp;�˱�֤����</label>	
					<input name="returnwarranty" style="vertical-align:top;width:100px" maxlength="15" value=<%=rs("returnwarranty")%> onKeyPress="javascript:CheckNum();" onKeyUp="this.value=this.value.replace(/[^\d.]/g,'')">
        </td>
      </tr>	        
      <tr>
        <td align="right" height="30" valign="top">�ٻ������</td>
        <td class="category" valign="top">
        	<textarea name="stockmiss" cols="30" rows="3" value=<%=rs("stockmiss")%>></textarea>&nbsp;&nbsp;&nbsp;
        	<label for="claiminformation" style="vertical-align:top">�������</label>
        	<textarea name="claiminformation" id="claiminformation" cols="30" rows="3" value=<%=rs("claiminformation")%>></textarea>&nbsp;&nbsp;&nbsp;
					<label for="claimstatus" style="vertical-align:top">����״̬</label>
					<%
					sql="select * from claimstatus"
					set rs_claimstatus=conn.execute(sql)
					%>
					<select name="claimstatus">
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
							<option value="������">������</option>
							<option value="������">������</option>
							<option value="��������">��������</option>
					</select>     -->
					<label for="claimprocessor" style="vertical-align:top">���⸺����</label>
					<select name="claimprocessor" style="vertical-align:top">
	    			<% for i = 0 to customCount-1 %>
	 						<option value=<%=custom(i)%> <%if rs("claimprocessor")=custom(i) then%>selected="selected"<%end if%>><%=custom(i)%></option>
	 					<% next %>  			
					</select>								
        </td>
      </tr>	       
      <tr>
		  <input type="submit" value=" ȷ���޸� " onClick="return check1();" class="button">&nbsp;&nbsp;&nbsp;
		  <input type="hidden" name="hid1" value="ok">
		  <input type="button" value=" �����޸ķ��� " onClick="if (confirm('ȷ��Ҫ�����޸���')) {window.open('../master/delete_lock_table.asp?tablename=shipment&combinedkey=<%=request("shipment")%>'); window.location.href='shipment.asp';}" class="button">
      </tr>    
</table>  
</form>
<table align="center" cellpadding="4" cellspacing="1" class="toptable grid" border="1">
	  <form name="formitem">
      <tr>
		<input type="button" onclick="addNewRow()" value="����һ��" class="button"/>&nbsp;&nbsp;&nbsp;
		<input type="button" onclick="removeRow()" value="ɾ��һ��" class="button"/>
      </tr>	      
      <tr>
		<table align="center" cellpadding="4" cellspacing="1" class="toptable grid" border="1" id="item">
		  <tr>
		  	<td width="20"></td>			
	        <td width="30" align="center">��Ŀ��</td>
	        <td width="100" align="center">Ʒ��</td>
	        <td width="100" align="center" >�ͻ�</td> 	
	        <td width=100" align="center" >���</td> 		
	        <td width=80" align="center" >��ͬ����</td> 			
	        <td width="40" align="center" >������λ</td> 		
	        <td width="80" align="center" >ʵ�ʾ���</td> 				    
	        <td width="80" align="center" >�ɹ��۸�</td> 			
	        <td width="40" align="center" >�۸�λ</td> 				        	            		        	        		        
	        <td width="50" align="center" >��������</td> 		
	        <td width="50" align="center" >����</td> 		
	        <td width="100" align="center" >��Ʊ�ܽ��</td> 		
	        <td width="50" align="center" >����</td> 		
	        <td width="100" align="center" >β����</td> 					        			        			        			        			        
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
	        <td width="40" align="center" >����</td> 		
	        <td width="80" align="center" ><input type='text' name="actualweight<%=rs_items("itemnum")%>" id="actualweight<%=rs_items("itemnum")%>" value='<%=rs_items("actualnetweight")%>' style='width:80px'/></td>
	        <td width="80" align="center" ><input type='text' name="purchaseprice<%=rs_items("itemnum")%>" id="purchaseprice<%=rs_items("itemnum")%>" value='<%=rs_items("purchaseprice")%>' style='width:80px'/></td> 	
	        <td width="40" align="center" >����</td> 				        	            		        	        		        
	        <td width="50" align="center" ><input name="producedate<%=rs_items("itemnum")%>" id="producedate<%=rs_items("itemnum")%>" value="<%=rs_items("productiondate")%>" readonly style='width:80px' onClick="JavaScript:window.open('../day.asp?field=producedate<%=rs_items("itemnum")%>&oldDate='+producedate<%=rs_items("itemnum")%>.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');"/></td> 		
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