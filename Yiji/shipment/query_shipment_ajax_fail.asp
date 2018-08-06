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
<title><%=dianming%> - 船期表查询</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script type="text/javascript" src="../js/jquery-3.2.1.js"></script>
<script type="text/javascript" src="../js/jquery-ui.js"></script>
<script type="text/javascript" src="../js/jquery.dataTables.min.js"></script>
<!--支持按钮功能-->
<script type="text/javascript" src="../js/dataTables.buttons.min.js"></script>
<script type="text/javascript" src="../js/buttons.colVis.min.js"></script>
<!--固定列-->
<script type="text/javascript" src="../js/dataTables.fixedColumns.min.js"></script>
<!--拖拽列-->
<script type="text/javascript" src="../js/dataTables.colReorder.min.js"></script>
<!--导出到excel-->
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
<br><center><img src="../images/note.gif" align="absmiddle">&nbsp;<font color="#FF0000">你不具备此权限，请与管理员联系！</font></center>
<%  
  response.end
end if
%>

<script>
//document.getElementById('queryresult').style.display='none';

  $(function(){

//日期控件
    $("#datefrom").datepicker();
    $("#dateto").datepicker();
//DataTable控件
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
                  text: "显示/隐藏列"
                  //columns: ":not(.noVis)"
              },
              {
                extend: "excel",
                text: "导出到Excel"
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

//table.column( "修改" ).order( false ).draw();    
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
	  win=window.open("../produit/print_customer_shipment.asp?shipmentnum="+shipment,"详细信息"); 	
		//win=window.open("../produit/print_customer_shipment.asp?shipmentnum=10000011","详细信息"); 	
	};
		
	//win=window.open("../produit/print_customer_shipment.asp?shipmentnum=10000011","Detail Information"); 
	win.focus();
}

//窗口表格增加一行
function addNewCriteria(){
   var tabObj=document.getElementById("search");//获取添加数据的表格
   var rowsNum = tabObj.rows.length;  //获取当前行数
   var colsNum=tabObj.rows[0].cells.length;//获取行的列数
   var myNewRow = tabObj.insertRow(rowsNum);//插入新行.
      
   var fieldObj=document.getElementsByName("fieldno");//取得所有行的fieldno
   var fieldNo=1;
   if (fieldObj.length==0) {
     fieldNo=1;//如果没有item,给1
     
  }else{
     fieldNo=parseInt(fieldObj[fieldObj.length-1].value) + 1; //取最大行fieldno加1
  } 

   var newTdObj1=myNewRow.insertCell(0);
   newTdObj1.innerHTML="<input type='hidden' name='fieldno' id='fieldno' value='"+fieldNo+"'/>";

   var newTdObj2=myNewRow.insertCell(1);
   newTdObj2.innerHTML="<select name='field"+fieldNo+"' id='field"+fieldNo+"' style='width:120px' align='left'>"
      +" <option value='buyer'>采购</option>"
      +" <option value='sales'>销售</option>"
      +" </select>";
   //newTdObj2.align="left";
   var newTdObj3=myNewRow.insertCell(2);
   newTdObj3.innerHTML="<input type='text' name='fieldvalue"+fieldNo+"' id='fieldvalue"+fieldNo+"' style='width:120px' align='left'/>";
   //newTdObj3.innerHTML="<input type='text' name='fieldvalue2' id='fieldvalue2'/>";
   //newTdObj3.align="left";
}

function queryresult()
 {
 var sql="select a.*, b.* from shipment a left join shipmentitem b on a.shipmentnum = b.shipmentnum where a.sales='张骥'"
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
'取得搜索关键字  
nowsales=2
'nowsales=request("fieldno").count
'sql="select a.*, b.* from shipment a left join shipmentitem b on a.shipmentnum = b.shipmentnum where a."
sql="select a.*, b.* from shipment a left join shipmentitem b on a.shipmentnum = b.shipmentnum where a.sales='张骥'"  '&request("fieldvalue1")&"'"
'sql="select a.*, b.* from shipment a left join shipmentitem b on a.shipmentnum = b.shipmentnum where a.sales='"&nowsales&"'"

%>
<table width="30%" border="0" cellpadding="0" cellspacing="0" align="left">
<form name="form2">
  <table>
    <tr> 
      <td align="right">
        <input type="button" value=" 添加一个查询条件"&<%=nowsales%> onclick="addNewCriteria()" class="button" > &nbsp;
      </td>  
      <td align="right">
        <input type="button" value=" 查询 " onclick="queryresult()" class="button">&nbsp;
        
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
    		     <option value="buyer">采购</option>
             <option value="sales">销售</option>
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