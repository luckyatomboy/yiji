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
<title><%=dianming%> - ���ڱ��ѯ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script type="text/javascript" src="../js/jquery-3.2.1.js"></script>
<script type="text/javascript" src="../js/jquery-ui.js"></script>
<script type="text/javascript" src="../js/jquery.dataTables.min.js"></script>
<!--֧�ְ�ť����-->
<script type="text/javascript" src="../js/dataTables.buttons.min.js"></script>
<script type="text/javascript" src="../js/buttons.colVis.min.js"></script>
<!--�̶���-->
<script type="text/javascript" src="../js/dataTables.fixedColumns.min.js"></script>
<!--��ק��-->
<script type="text/javascript" src="../js/dataTables.colReorder.min.js"></script>
<!--������excel-->
<script type="text/javascript" src="../js/buttons.flash.min.js"></script>
<script type="text/javascript" src="../js/jszip.min.js"></script>
<script type="text/javascript" src="../js/buttons.html5.min.js"></script>

<link href="../style/jquery-ui.css" rel="stylesheet" type="text/css">
<link href="../style/style.css" rel="stylesheet" type="text/css">
<link href="../style/jquery.dataTables.css" rel="stylesheet" type="text/css">
<link href="../style/buttons.dataTables.min.css" rel="stylesheet" type="text/css">
<link href="../style/fixedColumns.dataTables.min.css" rel="stylesheet" type="text/css">

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

<script>
//document.getElementById('queryresult').style.display='none';


    //$("#queryresult").style.display="inline";

//var table = $("#queryresult").DataTable();

//table.column( "�޸�" ).order( false ).draw();    
  $(function(){

//���ڿؼ�
    $("#datefrom").datepicker();
    $("#dateto").datepicker();
//DataTable�ؼ�
    //$("#queryresult").DataTable();
    //$("#queryresult").style.display="none";
    var table=$("#queryresult").DataTable( {
          dom: "Blfrtip",
          stateSave: true,
          colReorder: false,      
          buttons: [
              {
                  extend: "colvis",
                  collectionLayout: "fixed four-column",
                  columns: ".setvisible",
                  text: "��ʾ/������"
                  //columns: ":not(.noVis)"
              },
              {
                extend: "excel",
                text: "������Excel"
              }
          ],
          columnDefs: [
              { targets: [-1, -2], orderable: false },
              { targets: "ininovis", visible: false }
              //{ targets: "_all", visible: false}
          ]
   

} );
  });


function print_customer_shipment() { 
  var win;
  var shipment="";
  
	var chkObj=document.getElementsByName("sel");
	
	for(var k=0;k<chkObj.length;k++){
		if(chkObj[k].checked){
			if (shipment==""){
		 		shipment+=chkObj[k].value;
		 	}else{
		 		shipment=shipment+","+chkObj[k].value;
		 	}
		}
	}
  
	if (shipment != "") {
	  win=window.open("../produit/print_customer_shipment.asp?shipmentnum="+shipment,"��ϸ��Ϣ"); 	
		//win=window.open("../produit/print_customer_shipment.asp?shipmentnum=10000011","��ϸ��Ϣ"); 	
	};
		
	//win=window.open("../produit/print_customer_shipment.asp?shipmentnum=10000011","Detail Information"); 
	win.focus();
}

//���ڱ������һ��
function addNewCriteria(){
   var tabObj=document.getElementById("search");//��ȡ������ݵı��
   var rowsNum = tabObj.rows.length;  //��ȡ��ǰ����
   var myNewRow = tabObj.insertRow(rowsNum);//��������.
   var fieldNo=rowsNum+1;
   document.getElementById("fieldnum").value=fieldNo;

   var newTdObj1=myNewRow.insertCell(0);
   newTdObj1.innerHTML="<select name='fieldname"+fieldNo+"' id='fieldname"+fieldNo+"' style='width:120px' align='left'>"
      +" <option value='a.buyer' selected>�ɹ�</option>"
      +" <option value='a.sales' >����</option>"
      +" <option value='a.handler'>����</option>"
      +" <option value='a.country'>����</option>"
      +" <option value='a.plant'>����</option>"
      +" <option value='a.destination'>Ŀ�ĸ�</option>"
      +" <option value='a.terminal'>������ͷ</option>"
      +" <option value='a.carrier'>����˾</option>"
      +" <option value='a.shipname'>��������</option>"
      +" <option value='a.shipdate'>װ����</option>"
      +" <option value='a.boarddate'>������</option>"      
      +" <option value='b.material'>Ʒ��</option>"
      +" <option value='b.customer'>�ͻ�</option>"
      +" </select>";
   var newTdObj2=myNewRow.insertCell(1);
   newTdObj2.innerHTML="<select name='operation"+fieldNo+"' id='operation"+fieldNo+"' style='width:100px' align='left'>"
      +" <option value='=' selected>����</option>"
      +" <option value='<>' >������</option>"
      +" <option value='>='>���ڵ���</option>"
      +" <option value='<='>С�ڵ���</option>"
      +" </select>";      
   var newTdObj3=myNewRow.insertCell(2);
   newTdObj3.innerHTML="<input type='text' name='fieldvalue"+fieldNo+"' id='fieldvalue"+fieldNo+"' style='width:120px' align='left'/>";
}

</script>

<%
if request("query")="ok" then 
'ȡ�������ؼ���  
  sql="select a.*, b.itemnum, b.material, b.customer, b.spec, b.contractweight, b.contractweightuom, b.purchaseprice, b.purchasepriceunit, b.productiondate, b.casenumber, b.actualnetweight, b.invoiceamount, b.invoicecurrency, b.finalpayment, b.package from shipment a left join shipmentitem b on a.shipmentnum = b.shipmentnum where 1=1 "
  for i = 1 to request("fieldnum")
  'for i = 1 to 3 
    sql_ind="X"
    nowfieldno=i
    nowfield=request("fieldname"&i)
    nowoperation=request("operation"&i)
    nowfieldvalue=request("fieldvalue"&i)
    if nowfieldvalue<>"" then
      if instr(nowfield,"date") > 0 then '���������'
        nowchar = "#"
      else
        nowchar = "'"
      end if
      if nowoperation="=" then
        if nowchar="#" then
          sql=sql&" and "&nowfield&" = #"&nowfieldvalue&"#"
        else
          sql=sql&" and "&nowfield&" like '%"&nowfieldvalue&"%'"
        end if
      elseif nowoperation = "<>" then
        if nowchar="#" then
          sql=sql&" and "&nowfield&" <> #"&nowfieldvalue&"#"
        else
          sql=sql&" and "&nowfield&" not like '%"&nowfieldvalue&"%'"
        end if
      elseif nowoperation = ">=" then
        if nowchar="#" then
          sql=sql&" and "&nowfield&" >= #"&nowfieldvalue&"#"
        else
          sql=sql&" and "&nowfield&" >= '"&nowfieldvalue&"'"
        end if
      elseif nowoperation = "<=" then
        if nowchar="#" then
          sql=sql&" and "&nowfield&" <= #"&nowfieldvalue&"#"
        else
          sql=sql&" and "&nowfield&" <= '"&nowfieldvalue&"'"
        end if
      else
        sql=sql&" and "&nowfield&" like '%"&nowfieldvalue&"%'"
      end if    
    end if
  next

  'sql="select a.*, b.* from shipment a left join shipmentitem b on a.shipmentnum = b.shipmentnum where a.buyer like '��%'"

else

  nowfieldnum=request("fieldnum")

end if 

sqltemp="select a.*, b.* from shipment a left join shipmentitem b on a.shipmentnum = b.shipmentnum where a.buyer='����'"


'nowsales=request("fieldvalue1")
'nowsales=request("fieldno").count
'sql="select a.*, b.* from shipment a left join shipmentitem b on a.shipmentnum = b.shipmentnum where a."
'sql="select a.*, b.* from shipment a left join shipmentitem b on a.shipmentnum = b.shipmentnum where a.sales='����'"  '&request("fieldvalue1")&"'"
'sql="select a.*, b.* from shipment a left join shipmentitem b on a.shipmentnum = b.shipmentnum where a.sales='"&nowsales&"'"
%>
<form name="form1">
<table width="30%" border="0" cellpadding="0" cellspacing="0" align="left">

  <table>
    <tr> 
      <td align="right">
        <input type="button" value=" ���һ����ѯ����" onclick="addNewCriteria()" class="button" > &nbsp;
      </td>  
      <td align="right">
        <input type="submit" value=" ��ѯ " class="button">&nbsp;
        <input type="hidden" name="query" value="ok">
        <input type="hidden" name="fieldnum" id="fieldnum" value="1">
      </td>
    </tr>
  </table>
  <table id="search">
    <tr> 
    <!--
      <td align="middle" >
       <input type="text" name="fieldnum" id="fieldnum" value="1">
      </td>
    -->
    	<td align="left" >
      
    			<select name="fieldname1" id="fieldname1" style='width:120px' align='left'>
    		     <option value="a.buyer">�ɹ�</option>
             <option value="a.sales">����</option>
             <option value="a.handler">����</option>
             <option value="a.country">����</option>
             <option value="a.plant">����</option>
             <option value="a.destination">Ŀ�ĸ�</option>
             <option value="a.terminal">������ͷ</option>
             <option value="a.carrier">����˾</option>
             <option value="a.shipname">��������</option>
             <option value="a.shipdate">װ����</option>
             <option value="a.boarddate">������</option>             
             <option value="b.material">Ʒ��</option>
             <option value="b.customer">�ͻ�</option>
    		  </select>	  
      
    	</td>
      <td align="left" >
      
          <select name="operation1" id="operation1" style='width:100px' align='left'>
             <option value="=">����</option>
             <option value="<>">������</option>
             <option value=">=">���ڵ���</option>
             <option value="<=">С�ڵ���</option>
          </select>   
      
      </td>      
      <td align="left">
          <input type="text" name="fieldvalue1" id="fieldvalue1" style='width:120px' align='left'>
      </td>  
    </tr>
    
  </table>


</table>
</form>


<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../images/r_1.gif" alt="" /></td>
<td width="100%" background="../images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td width="10%" align="left">&nbsp;���ڱ��ѯ</td>
<!--	  <td align="right"><input type="button" value="��ӡ�ͻ����ڱ�" class="button" onClick="javascript:var win=window.open('../produit/print_customer_shipment.asp?shipmentnum=10000011','��ϸ��Ϣ'); win.focus()"></td>-->
	  <td align="right"><input type="button" value="��ӡ�ͻ����ڱ�" class="button" onClick="print_customer_shipment()"></td>
    </tr>
  </table>
</td>
<td><img src="../images/r_2.gif" alt="" /></td>
</tr>


<tr>
<td></td>
<td>
<!--startprint-->
<!--<table align="center" cellpadding="4" cellspacing="1" class="toptable grid" border="1" id="example">-->
<table width="100%" class="dataTable stripe" id="queryresult" cellspacing="1">
<!--
  <input type="hidden" name="queryform" value="<%=request("queryform")%>">
  <input type="hidden" name="form" value="<%=request("form")%>">
  <input type="hidden" name="field" value="<%=request("field")%>">
  <input type="hidden" name="field2" value="<%=request("field2")%>">
  <input type="hidden" name="field3" value="<%=request("field3")%>">
-->
  <thead class="sorting">
<!--Header-->
	<td class="category" width="100" height="30">���ڱ����</td>
	<td class="category setvisible" width="80" height="30">״̬</td>
	<td class="category setvisible" width="100" height="30">�ͻ�����</td>
	<td class="category setvisible ininovis" width="100" height="30">��Ӧ������</td>
	<td class="category setvisible ininovis" width="120" height="30">��ͬ��</td>
	<td class="category setvisible" width="80" height="30">��Ŀ��</td>	
	<td class="category setvisible" width="80" height="30">Ʒ��</td>
	<td class="category setvisible" width="80" height="30">����</td>
	<td class="category setvisible" width="80" height="30">����</td>

	<td class="category setvisible" width="80" height="30">������</td>
	<td class="category setvisible" width="80" height="30">��������</td>
	<td class="category setvisible" width="80" height="30">���</td>
  <td class="category setvisible" width="100" height="30">����˾��</td>
  <td class="category setvisible ininovis" width="100" height="30">���㵥״̬</td>
  <td class="category setvisible ininovis" width="80" height="30">ҵ������</td>
  <td class="category setvisible ininovis" width="100" height="30">����˾</td>
  <td class="category setvisible ininovis" width="50" height="30">�ɹ�</td>
  <td class="category setvisible ininovis" width="50" height="30">����</td>
  <td class="category setvisible ininovis" width="50" height="30">����</td>
  <td class="category setvisible ininovis" width="50" height="30">Ʒ��</td>
  <td class="category setvisible ininovis" width="60" height="30">��������</td>    

  <td class="category setvisible ininovis" width="100" height="30">�ۿ�</td>
  <td class="category setvisible ininovis" width="100" height="30">������ͷ</td>
  <td class="category setvisible ininovis" width="100" height="30">��������</td>

  <td class="category setvisible ininovis" width="100" height="30">��������</td>
  <td class="category setvisible ininovis" width="100" height="30">��֤�뱸��</td>
  <td class="category setvisible ininovis" width="100" height="30">����֤</td>
  <td class="category setvisible ininovis" width="100" height="30">�Զ�֤</td>
  <td class="category setvisible ininovis" width="100" height="30">�Զ�֤�ϱ���</td>
  <td class="category setvisible ininovis" width="100" height="30">�Զ�֤������</td>
  
  <td class="category setvisible ininovis" width="100" height="30">Ԥ����</td>
  <td class="category setvisible ininovis" width="100" height="30">������</td>
  <td class="category setvisible ininovis" width="100" height="30">����֧����</td>     
  <td class="category setvisible ininovis" width="100" height="30">����</td>    

  <td class="category setvisible ininovis" width="100" height="30">Ԥ��װ���·�</td>
  <td class="category setvisible ininovis" width="100" height="30">ʵ��װ����</td>
  
  <td class="category setvisible ininovis" width="100" height="30">�ͻ�������</td>    


  <td class="category setvisible ininovis" width="100" height="30">�ᵥ��</td>
  <td class="category setvisible ininovis" width="100" height="30">����֤��</td>     
  <td class="category setvisible ininovis" width="100" height="30">Ǧ���</td>     
  <td class="category setvisible ininovis" width="60" height="30">������Ϣ</td>
  <td class="category setvisible ininovis" width="60" height="30">����֤��</td>     
  <td class="category setvisible ininovis" width="50" height="30">��ǩ</td>     

  <td class="category setvisible ininovis" width="100" height="30">Ԥ��������</td>
  <td class="category setvisible ininovis" width="100" height="30">Ԥ������</td>
  <td class="category setvisible ininovis" width="100" height="30">����</td>     
  <td class="category setvisible ininovis" width="100" height="30">�����յ�����</td>     
  <td class="category setvisible ininovis" width="60" height="30">�ɽ�����</td>  

  <td class="category setvisible ininovis" width="100" height="30">β��֧������</td>
  <td class="category setvisible ininovis" width="100" height="30">������</td>
  <td class="category setvisible ininovis" width="100" height="30">��ҽ��</td>     
  <td class="category setvisible ininovis" width="100" height="30">��������</td>     

  <td class="category setvisible ininovis" width="100" height="30">��������</td>
  <td class="category setvisible ininovis" width="100" height="30">Ԥ���굥����</td>
  <td class="category setvisible ininovis" width="100" height="30">��˾�굥����</td>     
  <td class="category setvisible ininovis" width="100" height="30">�����굥����</td>      

  <td class="category setvisible ininovis" width="100" height="30">�ͻ�����</td>
  <td class="category setvisible ininovis" width="100" height="30">�ͻ�����</td>
  

  <td class="category setvisible ininovis" width="100" height="30">��֤��״̬</td>
  <td class="category setvisible ininovis" width="100" height="30">��֤����</td>
  <td class="category setvisible ininovis" width="100" height="30">�˱�֤����</td>     
  <td class="category setvisible ininovis" width="60" height="30">����״̬</td>     
  <td class="category setvisible ininovis" width="60" height="30">���⸺����</td>    

<!--Item-->
  <td class="category setvisible ininovis" width="100" height="30">���</td>
  <td class="category setvisible ininovis" width="100" height="30">��װ</td>
  <td class="category setvisible ininovis" width="100" height="30">��ͬ����</td>     
  <td class="category setvisible ininovis" width="60" height="30">������λ</td>     
  <td class="category setvisible ininovis" width="100" height="30">ʵ�ʾ���</td>    
  <td class="category setvisible ininovis" width="100" height="30">�ɹ��۸�</td>
  <td class="category setvisible ininovis" width="60" height="30">�۸�λ</td>
  <td class="category setvisible ininovis" width="100" height="30">��������</td>     
  <td class="category setvisible ininovis" width="100" height="30">����</td>     
  <td class="category setvisible ininovis" width="100" height="30">��Ʊ�ܽ��</td>   
  <td class="category setvisible ininovis" width="60" height="30">����</td>     
  <td class="category setvisible ininovis" width="100" height="30">β����</td> 


  <td class="category">�޸�</td>
  <td class="category">ѡ��</td>
  </thead>
  <!--</tr>-->
  <tbody>
  <%
  'if nowfieldno<>"" then
  if sql_ind<>"" then
  set rs_shipment =server.createobject("ADODB.RecordSet")	
  'rs_shipment.open sqltemp,conn,1,1
  rs_shipment.open sql,conn,1,1
  if not rs_shipment.eof then
  do while rs_shipment.eof=false
  %>

    <tr align="center" 
        <%if rs_shipment("status")="����" then%>bgcolor="darkgrey"<%elseif rs_shipment("status")="ͨ����" then%>bgcolor="lawngreen"<%elseif rs_shipment("status")="���ͻ�" then%>bgcolor="red"<%end if%>>     

<!--Header-->
  <td align="center" height="30"><%=rs_shipment("shipmentnum")%></td>    
  <td align="center" class="setvisible"><%=rs_shipment("status")%></td>	  
  <td align="center" class="setvisible"><%=rs_shipment("customer")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("vendor")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("contract")%></td>
  <td align="center" class="setvisible"><%=rs_shipment("itemnum")%></td>  
  <td align="center" class="setvisible"><%=rs_shipment("material")%></td>
  <td align="center" class="setvisible"><%=rs_shipment("country")%></td>
  <td align="center" class="setvisible"><%=rs_shipment("plant")%></td>

  <td align="center" class="setvisible"><%=rs_shipment("boarddate")%></td>
  <td align="center" class="setvisible"><%=rs_shipment("deliverydate")%></td>
  <td align="center" class="setvisible"><%=rs_shipment("case")%></td>
  <td align="center" class="setvisible"><%=rs_shipment("carrier")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("invoicestatus")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("trantype")%></td>  
  <td align="center" class="setvisible ininovis"><%=rs_shipment("agent")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("buyer")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("sales")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("handler")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("materialtype")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("incoterm")%></td>  

  <td align="center" class="setvisible ininovis"><%=rs_shipment("destination")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("terminal")%></td>  
  <td align="center" class="setvisible ininovis"><%=rs_shipment("shipname")%></td>  

  <td align="center" class="setvisible ininovis"><%=rs_shipment("piwen")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("twodocumentreadydate")%></td>  
  <td align="center" class="setvisible ininovis"><%=rs_shipment("dongjian")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("zidong")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("zidongapplydate")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("zidongreportdate")%></td>

  <td align="center" class="setvisible ininovis"><%=rs_shipment("planinsurancedate")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("supinsurancedate")%></td>    
  <td align="center" class="setvisible ininovis"><%=rs_shipment("insurancepaymentdate")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("insurancenumber")%></td>    

  <td align="center" class="setvisible ininovis"><%=rs_shipment("planship")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("shipdate")%></td>    

  <td align="center" class="setvisible ininovis"><%=rs_shipment("customerdeliverydate")%></td>   


  <td align="center" class="setvisible ininovis"><%=rs_shipment("ladnumber")%></td>    
  <td align="center" class="setvisible ininovis"><%=rs_shipment("weishengzheng")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("locknumber")%></td>   
  <td align="center" class="setvisible ininovis"><%if rs_shipment("einformation")="True" then%>��<%else%>��<%end if%></td>
  <td align="center" class="setvisible ininovis"><%if rs_shipment("MuslimCertification")="True" then%>��<%else%>��<%end if%></td>    
  <td align="center" class="setvisible ininovis"><%if rs_shipment("elabel")="True" then%>��<%else%>��<%end if%></td>

  <td align="center" class="setvisible ininovis"><%=rs_shipment("prepaydate")%></td>  
  <td align="center" class="setvisible ininovis"><%=rs_shipment("prepayment")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("trancurrency")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("downpaymentreceiptdate")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("tradingterm")%></td>  

  <td align="center" class="setvisible ininovis"><%=rs_shipment("finalpaymentdate")%></td>  
  <td align="center" class="setvisible ininovis"><%=rs_shipment("freestayperiod")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("shouyi")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("passdate")%></td>  

  <td align="center" class="setvisible ininovis"><%=rs_shipment("documentarrivaldate")%></td>  
  <td align="center" class="setvisible ininovis"><%=rs_shipment("planretirebilldate")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("internalretirebilldate")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("externalretirebilldate")%></td>    

  <td align="center" class="setvisible ininovis"><%=rs_shipment("cargodate")%></td>  
  <td align="center" class="setvisible ininovis"><%=rs_shipment("cargodirection")%></td>
  
  <td align="center" class="setvisible ininovis"><%=rs_shipment("warrantystatus")%></td>  
  <td align="center" class="setvisible ininovis"><%=rs_shipment("warranty")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("returnwarranty")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("claimstatus")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("claimprocessor")%></td>  

<!--Item-->
  <td align="center" class="setvisible ininovis"><%=rs_shipment("spec")%></td>  
  <td align="center" class="setvisible ininovis"><%=rs_shipment("package")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("contractweight")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("contractweightuom")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("actualnetweight")%></td>  
  <td align="center" class="setvisible ininovis"><%=rs_shipment("purchaseprice")%></td>  
  <td align="center" class="setvisible ininovis"><%=rs_shipment("purchasepriceunit")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("productiondate")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("casenumber")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("invoiceamount")%></td>    
  <td align="center" class="setvisible ininovis"><%=rs_shipment("invoicecurrency")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("finalpayment")%></td>     


  <td align="center"> 
    	<a href="change_shipment_new.asp?form=<%=request("form")%>&shipment=<%=rs_shipment("shipmentnum")%>&keyword=<%=nowkeyword%>" target="_blank"><img src="../images/res.gif" border="0" hspace="2" align="absmiddle">�޸�</a>
  </td>
  <td align="center">
  	<input type="checkbox" name="sel" value="<%=rs_shipment("shipmentnum")&rs_shipment("itemnum")%>" style="border:0">
  </td>
<!--	<td align="center"><input type="checkbox" name='chkArr'  id='chkArr' />";  -->
  </tr>
  <%
  	'end if
    rs_shipment.movenext
  loop
  end if
  end if
  %>
  </tbody>


</table>
<!--endprint-->
</td>
<td></td>
</tr>

</table>

</body>
</html>