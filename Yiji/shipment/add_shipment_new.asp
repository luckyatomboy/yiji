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
<title><%=dianming%> - ���ڱ�</title>
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
<br><center><img src="../images/note.gif" align="absmiddle">&nbsp;<font color="#FF0000">�㲻�߱���Ȩ�ޣ��������Ա��ϵ��</font></center>
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
alert("������ĺ�ͬ���Ѿ����ڣ����������룡")
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
if request("dongjian")="����ѡ�񶯼����֤" then
	nowdongjian=""
else
	nowdongjian=request("dongjian")
end if
nowzidongcom=request("zidongcom")
if request("zidong")="����ѡ���Զ����֤" then
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


if request("trantype")="����" then
	sql="select top 1 * from shipment where shipmentnum like '1%' order by shipmentnum desc"
else
	sql="select top 1 * from shipment where shipmentnum like '2%' order by shipmentnum desc"
end if
set rs_count=conn.execute(sql)

if rs_count.bof and rs_count.eof then
	if request("trantype")="����" then
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

'sql="insert into shipmentitem(shipmentnum,itemnum,material,customer,spec,contractweight,contractweightuom,purchaseprice,purchasepriceunit,productiondate,casenumber,actualnetweight,actualnetweightuom,invoiceamount,invoicecurrency,finalpayment,finalpaymentcurrency) values("&nowshipment&",'"&nowitemno&"','"&nowmatitem&"','"&nowcusitem&"','"&nowspecitem&"','"&nowcontweightitem&"','����','"&nowpurpriceitem&"','����','"&nowproducedateitem&"',"&nowcasenumitem&",'"&nowactualweightitem&"','����','"&nowinvamtitem&"','"&nowcurritem&"','"&nowfinalamtitem&"','"&nowcurritem&"')"
sql="insert into shipmentitem(shipmentnum,itemnum,material,customer,spec,contractweight,contractweightuom,actualnetweight,actualnetweightuom,purchaseprice,purchasepriceunit,productiondate,casenumber,invoiceamount,invoicecurrency,finalpayment,finalpaymentcurrency) values("&nowshipment&",'"&nowitemno&"','"&nowmatitem&"','"&nowcusitem&"','"&nowspecitem&"','"&nowcontweightitem&"','����','"&nowactualweightitem&"','����','"&nowpurpriceitem&"','����',#"&nowproducedateitem&"#,'"&nowcasenumitem&"',"&nowinvamtitem&",'"&nowcurritem&"',"&nowfinalamtitem&",'"&nowcurritem&"')"
conn.execute(sql)

next

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
   newTdObj7.innerHTML="����";   
   var newTdObj8=myNewRow.insertCell(7);
   newTdObj8.innerHTML="<input type='text' name='actualWeight"+itemNo+"' id='actualWeight"+itemNo+"' style='width:80px' value=0 onKeyPress=\""+"javascript:CheckNum();\""+" onKeyUp=\""+"this.value=this.value.replace(/[^\\d.]/g,'')\""+">";   		 
   var newTdObj9=myNewRow.insertCell(8);
   newTdObj9.innerHTML="<input type='text' name='purchasePrice"+itemNo+"' id='purchasePrice"+itemNo+"' style='width:80px' value=0 onKeyPress=\""+"javascript:CheckNum();\""+" onKeyUp=\""+"this.value=this.value.replace(/[^\\d.]/g,'')\""+">";   		 
   var newTdObj10=myNewRow.insertCell(9);
   newTdObj10.innerHTML="����";     
   var newTdObj11=myNewRow.insertCell(10);
   newTdObj11.innerHTML="<input name='produceDate"+itemNo+"' id='produceDate"+itemNo+"' readonly style='width:80px'"
//   	+ " <img src='../images/date.gif' align='absmiddle' style='cursor:pointer;'" 
   	+ " onClick=\""+"JavaScript:window.open('../day.asp?field=produceDate"+itemNo+"&oldDate=produceDate"+itemNo+".value','','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');\"/>";     
   var newTdObj12=myNewRow.insertCell(11);   	//����
	 newTdObj12.innerHTML="<input name='casenum"+itemNo+"' id='casenum"+itemNo+"' style='width:80px' onKeyPress=\""+"javascript:CheckNum();\""+" onKeyUp=\""+"this.value=this.value.replace(/[^\\d.]/g,'')\""+">";   	
   var newTdObj13=myNewRow.insertCell(12);   	//��Ʊ�ܽ��
	 newTdObj13.innerHTML="<input name='invAmount"+itemNo+"' id='invAmt"+itemNo+"' style='width:80px' value=0 onKeyPress=\""+"javascript:CheckNum();\""+" onKeyUp=\""+"this.value=this.value.replace(/[^\\d.]/g,'')\""+">";   		 
   var newTdObj14=myNewRow.insertCell(13);
   newTdObj14.innerHTML="<select name='invCurr"+itemNo+"' id='invCurr"+itemNo+"' style='width:50px'>"
	 		+" <option value='��Ԫ'>��Ԫ</option>"
	 		+" <option value='�ı�'>�ı�</option>"
	 		+" </select>";   
   var newTdObj15=myNewRow.insertCell(14);   	//β����
	 newTdObj15.innerHTML="<input name='finalAmount"+itemNo+"' id='finalAmt"+itemNo+"' style='width:80px' value=0 onKeyPress=\""+"javascript:CheckNum();\""+" onKeyUp=\""+"this.value=this.value.replace(/[^\\d.]/g,'')\""+">";   		 

  }

//���ڱ��ɾ��һ��
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
      <td>&nbsp;���ڱ� </td>
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
        <td width="20%" align="right" height="30">һ����Ϣ��</td>
        <td width="80%" class="category">
        	���ͣ�
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
					&nbsp;&nbsp;&nbsp;&nbsp;״̬
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
					&nbsp;&nbsp;&nbsp;&nbsp;���㵥״̬
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
	    	<td align="right" height="30">�ɹ���</td>
        <td class="category">		
					<select name="buyer">
	    			<% for i = 0 to buyerCount-1 %>
	 						<option value=<%=buyer(i)%>><%=buyer(i)%></option>
	 					<% next %>  			
					</select>						
					&nbsp;&nbsp;&nbsp;&nbsp;����	      
					<select name="custom">
	    			<% for i = 0 to customCount-1 %>
	 						<option value=<%=custom(i)%>><%=custom(i)%></option>
	 					<% next %>  			
					</select>				
					&nbsp;&nbsp;&nbsp;&nbsp;����	      
					<select name="sales">
	    			<% for i = 0 to salesCount-1 %>
	 						<option value=<%=sales(i)%>><%=sales(i)%></option>
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
							<option value="<%=rs_vendor("vendorname")%>"><%=rs_vendor("vendorname")%></option>
						<%
							rs_vendor.movenext
							loop
							rs_vendor.close
						%>
					</select>
					&nbsp;<font color="#ff0000">*</font>
					&nbsp;&nbsp;&nbsp;&nbsp;����				
					<select name="country">			
						<option selected="selected">---��ѡ��---</option>		  			
					</select>						
					&nbsp;&nbsp;&nbsp;&nbsp;����						
					<select name="plant"> <!-- onchange="this.selectedIndex=this.defaultIndex;">-->
						  <option selected="selected">---��ѡ��---</option>		
					</select>					
		  		&nbsp;&nbsp;&nbsp;���ʽ
<!--		  		<input name="incoterm" style="width:150px">							 -->
					<select name="incoterm">			
						  <option selected="selected">---��ѡ��---</option>										
					</select>				  				
				</td>
      </tr>
      <tr>	  
	    	<td align="right" height="30">��ͬ�ţ�</td>
        <td class="category">
		  		<input name="contract" value="<%=contract%>" style="width:150px" maxlength="20">
		  		&nbsp;<font color="#ff0000">*</font>
		  		&nbsp;&nbsp;&nbsp;&nbsp;Ʒ��
					<select name="materialtype">			
						  <option value="��">��</option>		
						  <option value="ţ">ţ</option>		
						  <option value="��">��</option>		
						  <option value="��">��</option>										
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
							<option value="<%=rs_agent("company")%>"><%=rs_agent("company")%></option>
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
							<option value="<%=rs_piwen("company")%>"><%=rs_piwen("company")%></option>
						<%
							rs_piwen.movenext
							loop
							rs_piwen.close
						%>								
					</select>			
					&nbsp;&nbsp;&nbsp;��֤
		  		<input name="twodocumentsready" id="twodocumentsready" readonly style="width:80px">
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=twodocumentsready&oldDate='+twodocumentsready.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">										 		
				</td>
      </tr>    
      <tr>	  
	    	<td align="right" height="30">�ۿڣ�</td>
        <td class="category">
					<select name="destination">			
						  <option value="�Ϻ�">�Ϻ�</option>		
						  <option value="����">����</option>		
						  <option value="���">���</option>		
						  <option value="����">����</option>										
					</select>		
					&nbsp;&nbsp;&nbsp;&nbsp;������ͷ
					<select name="terminal">			
						  <option value="������">������</option>		
						  <option value="����">����</option>		
						  <option value="��һ">��һ</option>		
						  <option value="����">����</option>										
					</select>			
					&nbsp;&nbsp;&nbsp;&nbsp;����˾
					<select name="carrier">			
						  <option value="ANL">ANL</option>		
						  <option value="MSC">MSC</option>		
						  <option value="CMA CGM">CMA CGM</option>		
						  <option value="COSCO">COSCO</option>										
					</select>				
					&nbsp;&nbsp;&nbsp;&nbsp;��������
					<input name="shipname" style="width:150px" maxlength="20">							
				</td>
      </tr>               
      <tr>	  
	    	<td align="right" height="30">����֤��</td>
        <td class="category">
		  		<input name="dongjian" readonly onClick="JavaScript:window.open('../huiyuan/query_license.asp?queryform=form1&licensetype=����֤&field=dongjian&field2=dongjiancom','','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=1200,height=470,top=100,left=20');" style="cursor:hand;width:120px" value="����ѡ�񶯼����֤">
		  		<input name="dongjiancom" readonly >
				</td>
      </tr>
      <tr>	  
	    	<td align="right" height="30">�Զ�֤��</td>
        <td class="category">
		  		<input name="zidong" readonly onClick="JavaScript:window.open('../huiyuan/query_license.asp?queryform=form1&licensetype=�Զ�֤&field=zidong&field2=zidongcom','','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=1200,height=470,top=100,left=20');" style="cursor:hand;width:120px" value="����ѡ���Զ����֤">
		  		<input name="zidongcom" readonly >
		  		&nbsp;&nbsp;&nbsp;&nbsp;�Զ�֤��������
		  		<input name="zidongapplydate" id="zidongapplydate" readonly style="width:80px">
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=zidongapplydate&oldDate='+zidongapplydate.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">		  				  		
		  		&nbsp;&nbsp;&nbsp;&nbsp;�Զ�֤�ϱ�����
		  		<input name="zidongreportdate" id="zidongreportdate" readonly style="width:80px">
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=zidongreportdate&oldDate='+zidongreportdate.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">		  				  		
				</td>
      </tr>
      <tr>	  
	    	<td align="right" height="30">Ԥ���գ�</td>
        <td class="category">
		  		<input name="planinsurance" id="planinsurance" style="width:80px">
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=planinsurance&oldDate='+planinsurance.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">
		  		&nbsp;&nbsp;&nbsp;&nbsp;������
		  		<input name="supinsurance" id="supinsurance" readonly style="width:80px">
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=supinsurance&oldDate='+supinsurance.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">
		  		&nbsp;&nbsp;&nbsp;&nbsp;����֧����
		  		<input name="insurancepayment" id="insurancepayment" readonly style="width:80px">
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=insurancepayment&oldDate='+insurancepayment.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">		  		
		  		&nbsp;&nbsp;&nbsp;&nbsp;����
		  		<input name="insurancenumber" style="width:120px">
				</td>
      </tr>      
      <tr>	  
	    	<td align="right" height="30">Ԥ��װ���·ݣ�</td>
        <td class="category">
		  		<input name="planship" style="width:100px">
		  		&nbsp;&nbsp;&nbsp;&nbsp;ʵ��װ����
		  		<input name="shipdate" id="shipdate" readonly style="width:80px">
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=shipdate&oldDate='+shipdate.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">
		  		&nbsp;&nbsp;&nbsp;&nbsp;Ԥ�Ƶ�����
		  		<input name="boarddate" id="boarddate" readonly style="width:80px">
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=boarddate&oldDate='+boarddate.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">		  		
		  		&nbsp;&nbsp;&nbsp;&nbsp;�ͻ�������
		  		<input name="plandeliverydate" style="width:100px">
				</td>
      </tr>
      <tr>	  
	    	<td align="right" height="30">��ţ�</td>
        <td class="category">
		  		<input name="case" style="width:120px" maxlength="15">
		  		&nbsp;&nbsp;&nbsp;�ᵥ��
		  		<input name="ladnumber" style="width:120px" maxlength="20">
		  		&nbsp;&nbsp;&nbsp;����֤��
		  		<input name="weishengzheng" style="width:120px" maxlength="20">	
					&nbsp;&nbsp;&nbsp;Ǧ���
					<input name="locknumber" style="width:120px" maxlength="20">		
		  		&nbsp;&nbsp;&nbsp;������Ϣ
					<select name="einformation">			
						  <option value="��">��</option>		
						  <option value="��">��</option>										
					</select>	
		  		&nbsp;&nbsp;&nbsp;����֤��
					<select name="MuslimCertification">			
						  <option value="��">��</option>		
						  <option value="��">��</option>										
					</select>											  		
				</td>
      </tr>
      <tr>	  
	    	<td align="right" height="30">Ԥ�������ڣ�</td>
        <td class="category">
		  		<input name="paydate" id="paydate" readonly style="width:80px">
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=paydate&oldDate='+paydate.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">
		  		&nbsp;&nbsp;&nbsp;Ԥ������
		  		<input name="payment" style="width:100px" maxlength="15" value=0 onKeyPress="javascript:CheckNum();" onKeyUp="this.value=this.value.replace(/[^\d.]/g,'')">
		  		&nbsp;&nbsp;&nbsp;����
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
				&nbsp;&nbsp;&nbsp;�����յ�����
				<input name="downpaymentdate" id="downpaymentdate" readonly style="width:80px">	
				<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=downpaymentdate&oldDate='+downpaymentdate.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">
				&nbsp;&nbsp;&nbsp;�ɽ�����
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
	    	<td align="right" height="30">β��֧�����ڣ�</td>
        <td class="category">
		  		<input name="finalpaymentdate" id="finalpaymentdate" readonly style="width:80px">
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=finalpaymentdate&oldDate='+finalpaymentdate.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">
		  		&nbsp;&nbsp;&nbsp;������
		  		<input name="freestayperiod" style="width:100px">
		  		&nbsp;&nbsp;&nbsp;����֤�汾
					<input name="weishengzhengversion" style="width:100px">
					&nbsp;&nbsp;&nbsp;��������
		  		<input name="passdate" id="passdate" readonly style="width:80px">
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=passdate&oldDate='+passdate.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">
				</td>
      </tr>       
      <tr>	  
	    	<td align="right" height="30">�������ڣ�</td>
        <td class="category">
		  		<input name="documentarrival" id="documentarrival" readonly style="width:80px">
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=documentarrival&oldDate='+documentarrival.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">
		  		&nbsp;&nbsp;&nbsp;Ԥ���굥����
		  		<input name="planretirebill" id="planretirebill" readonly style="width:80px">
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=planretirebill&oldDate='+planretirebill.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">		  		
		  		&nbsp;&nbsp;&nbsp;��˾�굥����
		  		<input name="internalretirebill" id="internalretirebill" readonly style="width:80px">
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=internalretirebill&oldDate='+internalretirebill.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">		  		
		  		&nbsp;&nbsp;&nbsp;�����굥����
		  		<input name="externalretirebill" id="externalretirebill" readonly style="width:80px">
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=externalretirebill&oldDate='+externalretirebill.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">		  		
				</td>
      </tr>       
      <tr>	  
	    	<td align="right" height="30">�ͻ����ڣ�</td>
        <td class="category">
		  		<input name="cargodate" id="cargodate" readonly style="width:80px">
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=cargodate&oldDate='+cargodate.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">
		  		&nbsp;&nbsp;&nbsp;�ͻ�����
		  		<input name="cargodir" style="width:150px">
		  		&nbsp;&nbsp;&nbsp;��������
		  		<input name="deliverydate" id="deliverydate" readonly style="width:80px">
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=deliverydate&oldDate='+deliverydate.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">		  		
				</td>
      </tr> 
      <tr>
        <td align="right" height="30" valign="top">��֤���⣺</td>
        <td class="category" valign="top">
        	<textarea name="documentissue" cols="30" rows="3"></textarea>&nbsp;&nbsp;&nbsp;
        	<label for="documentcheck" style="vertical-align:top"> ����� </label>
        	<textarea name="documentcheck" id="documentcheck" cols="30" rows="3"></textarea>&nbsp;&nbsp;&nbsp;
        	<label for="documentdelivery" style="vertical-align:top"> �ĵ���� </label>
        	<textarea name="documentdelivery" id="documentdelivery" cols="30" rows="3"></textarea>&nbsp;&nbsp;&nbsp;
        	<label for="clearcustom" style="vertical-align:top">��ؽ���</label>
        	<textarea name="clearcustom" id="clearcustom" cols="30" rows="3"></textarea>            	   	
        </td>
      </tr>	 
      <tr>
        <td align="right" height="30" valign="top">��֤�������</td>
        <td class="category" valign="top">
        	<textarea name="warrantyinformation" cols="30" rows="3"></textarea>&nbsp;&nbsp;&nbsp;
        	<label for="warrantystatus" style="vertical-align:top">��֤�����</label>
					<select name="warrantystatus" style="vertical-align:top">
	 						<option value="�����">�����</option>	
	 						<option value="���˱�">���˱�</option>	
	 						<option value="�ѵ���">�ѵ���</option>	
					</select>			
					<label for="warranty" style="vertical-align:top">&nbsp;&nbsp;&nbsp;��֤����</label>	
					<input name="warranty" style="vertical-align:top;width:100px" maxlength="15" value=0 onKeyPress="javascript:CheckNum();" onKeyUp="this.value=this.value.replace(/[^\d.]/g,'')">
					<label for="returnwarranty" style="vertical-align:top">&nbsp;&nbsp;&nbsp;�˱�֤����</label>	
					<input name="returnwarranty" style="vertical-align:top;width:100px" maxlength="15" value=0 onKeyPress="javascript:CheckNum();" onKeyUp="this.value=this.value.replace(/[^\d.]/g,'')">
        </td>
      </tr>	        
      <tr>
        <td align="right" height="30" valign="top">�ٻ������</td>
        <td class="category" valign="top">
        	<textarea name="stockmiss" cols="30" rows="3"></textarea>&nbsp;&nbsp;&nbsp;
        	<label for="claiminformation" style="vertical-align:top">�������</label>
        	<textarea name="claiminformation" id="claiminformation" cols="30" rows="3"></textarea>&nbsp;&nbsp;&nbsp;
					<label for="claimstatus" style="vertical-align:top">����״̬</label>
					<select name="claimstatus" style="vertical-align:top">
							<option value="������">������</option>
							<option value="������">������</option>
							<option value="��������">��������</option>
					</select>     
					<label for="claimprocessor" style="vertical-align:top">���⸺����</label>
					<select name="custom" style="vertical-align:top">
	    			<% for i = 0 to customCount-1 %>
	 						<option value=<%=custom(i)%>><%=custom(i)%></option>
	 					<% next %>  			
					</select>								
        </td>
      </tr>	       
      <tr>
		  <input type="submit" value=" ȷ��¼�� " onClick="return check1();" class="button">&nbsp;&nbsp;&nbsp;
		  <input type="hidden" name="hid1" value="ok">
		  <input type="reset" value=" ������д " class="button">
      </tr>    
</table>  
<table align="center" cellpadding="4" cellspacing="1" class="toptable grid" border="1">
<!--	  <form name="formitem" action="" method="post">-->
      <tr>
					<input type="button" onclick="addNewRow()" value="����һ��" class="button"/>&nbsp;&nbsp;&nbsp;
					<input type="button" onclick="removeRow()" value="ɾ��һ��" class="button"/>
      </tr>	      
      <tr>
					<table align="center" cellpadding="4" cellspacing="1" class="toptable grid" border="1" id="item">
					  <tr>
					  	<td width="20"></td>			
			        <td width="30" align="center" >��Ŀ��</td>
			        <td width="100" align="center" >Ʒ��</td>
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