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
    //$("#queryresult").style.display="inline";

//var table = $("#queryresult").DataTable();

//table.column( "�޸�" ).order( false ).draw();    
  });

</script>

<script>
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
   var colsNum=tabObj.rows[0].cells.length;//��ȡ�е�����
   var myNewRow = tabObj.insertRow(rowsNum);//��������.
      
   var fieldObj=document.getElementsByName("fieldno");//ȡ�������е�fieldno
   var fieldNo=1;
   if (fieldObj.length==0) {
     fieldNo=1;//���û��item,��1
     
  }else{
     fieldNo=parseInt(fieldObj[fieldObj.length-1].value) + 1; //ȡ�����fieldno��1
  } 

   var newTdObj1=myNewRow.insertCell(0);
   newTdObj1.innerHTML="<input type='hidden' name='fieldno' id='fieldno' value='"+fieldNo+"'/>";

   var newTdObj2=myNewRow.insertCell(1);
   newTdObj2.innerHTML="<select name='field"+fieldNo+"' id='field"+fieldNo+"' style='width:120px' align='left'>"
      +" <option value='buyer'>�ɹ�</option>"
      +" <option value='sales'>����</option>"
      +" </select>";
   //newTdObj2.align="left";
   var newTdObj3=myNewRow.insertCell(2);
   newTdObj3.innerHTML="<input type='text' name='fieldvalue"+fieldNo+"' id='fieldvalue"+fieldNo+"' style='width:120px' align='left'/>";
   //newTdObj3.innerHTML="<input type='text' name='fieldvalue2' id='fieldvalue2'/>";
   //newTdObj3.align="left";
}

function queryresult()
 {
 var sql="select a.*, b.* from shipment a left join shipmentitem b on a.shipmentnum = b.shipmentnum where a.sales='����'"
 if (window.XMLHttpRequest)
 {// code for IE7+, Firefox, Chrome, Opera, Safari
 xmlhttp=new XMLHttpRequest();
 }
 else
 {// code for IE6, IE5
 xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
 }
 xmlhttp.onreadystatechange=function()
 {
 if (xmlhttp.readyState==4 && xmlhttp.status==200)
 {
 document.getElementById("result").innerHTML=xmlhttp.responseText;
 }
 }
 xmlhttp.open("GET","query_result_simple.asp?sql="+sql,true);
 xmlhttp.send();
 }

</script>

<%
'ȡ�������ؼ���  
nowsales=2
'nowsales=request("fieldno").count
'sql="select a.*, b.* from shipment a left join shipmentitem b on a.shipmentnum = b.shipmentnum where a."
sql="select a.*, b.* from shipment a left join shipmentitem b on a.shipmentnum = b.shipmentnum where a.sales='����'"  '&request("fieldvalue1")&"'"
'sql="select a.*, b.* from shipment a left join shipmentitem b on a.shipmentnum = b.shipmentnum where a.sales='"&nowsales&"'"

%>
<table width="30%" border="0" cellpadding="0" cellspacing="0" align="left">
<form name="form2">
  <table>
    <tr> 
      <td align="right">
        <input type="button" value=" ���һ����ѯ����"&<%=nowsales%> onclick="addNewCriteria()" class="button" > &nbsp;
      </td>  
      <td align="right">
        <input type="button" value=" ��ѯ " onclick="queryresult()" class="button">&nbsp;
        
      </td>
    </tr>
  </table>  
  <table id="search">
    <tr> 
      <td align="middle" >
       <input type="hidden" name="fieldno" id="fieldno" value="1">
      </td>
    	<td align="left" >
      
    			<select name="field1" id="field1">
    		     <option value="buyer">�ɹ�</option>
             <option value="sales">����</option>
    		  </select>	  
      
    	</td>
      <td align="left">
          <input type="text" name="fieldvalue1" id="fieldvalue1">
      </td>  
    </tr>
  </table>

</form>  
</table>


<p><span id="result"></span></p>

</body>
</html>