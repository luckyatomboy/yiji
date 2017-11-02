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
  $(function(){
//日期控件
    $("#datefrom").datepicker();
    $("#dateto").datepicker();
//DataTable控件
    //$("#queryresult").DataTable();
    var table=$("#queryresult").DataTable( {
          dom: "Blfrtip",
          stateSave: true,
          colReorder: true,      
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
</script>

<%
'取得搜索关键字  
nowkeyword=request("keyword") 
nowsales=request("sales")
nowcustomer=request("customer")
nowfrom=request("datefrom")
nowto=request("dateto")
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" align="center">
<form name="form2">
  <input type="hidden" name="form" value="<%=request("form")%>">
  <input type="hidden" name="field" value="<%=request("field")%>">
  <input type="hidden" name="field2" value="<%=request("field2")%>">
  <input type="hidden" name="field3" value="<%=request("field3")%>">
<!--
  <input type="hidden" name="field4" value="<%=request("field4")%>">  
  <input type="hidden" name="field5" value="<%=request("field5")%>">
  <input type="hidden" name="field6" value="<%=request("field6")%>">
  <input type="hidden" name="field7" value="<%=request("field7")%>">
  <input type="hidden" name="field8" value="<%=request("field8")%>">    
  <input type="hidden" name="field9" value="<%=request("field9")%>">    
-->
  <input type="hidden" name="span1" value="<%=request("span1")%>">
  <tr> 
<!--  <td width="700" height="21" align="right"><font size=4><b>查询船期表</b></font></td> -->
	<td align="right">
	  搜索：
		<%
			'sql="select * from sales order by name"
      sql="select * from login where issales=1"
			set rs_sales=conn.execute(sql)
		%>   
			<select name="sales" onChange="form2.submit()">
		     <option value="">所有销售员</option>
	  <%
			do while rs_sales.eof=false
		%>
		    <option value="<%=rs_sales("username")%>"<%if trim(cstr(rs_sales("username")))=nowsales then%> selected="selected"<%end if%>><%=rs_sales("username")%></option>
		<%
			  rs_sales.movenext
			loop
		%>
		  </select>	  

		<%
			sql="select * from customer order by customername"
			set rs_customer=conn.execute(sql)
		%>
			<select name="customer" onChange="form2.submit()">
		     <option value="">所有客户</option>
	  <%
			do while rs_customer.eof=false
		%>
		    <option value="<%=rs_customer("customername")%>"<%if trim(cstr(rs_customer("customername")))=nowcustomer then%> selected="selected"<%end if%>><%=rs_customer("customername")%></option>
		<%
			  rs_customer.movenext
			loop
		%>
		  </select>	  	  
		<input type="text" name="keyword" placeholder="请输入船期表号码" size="20" value="<%=nowkeyword%>">
<!--	  <input type="submit" value=" 查询 " class="button">&nbsp; -->
	</td>
  </tr>
  <tr> 
	<td align="right">
	  到港日期：从
    <input name="datefrom" style="width:80px" id="datefrom">
	  &nbsp;&nbsp;&nbsp;&nbsp;到
    <input name="dateto" style="width:80px" id="dateto">
	  <input type="submit" value=" 查询 " class="button">&nbsp;
	</td>
  </tr>  
</form>  
</table>
<%
'  sql="select * from shipment where shipmentnum<=99999999"
  sql="select b.*, a.* from shipment a left join shipmentitem b on a.shipmentnum = b.shipmentnum where a.shipmentnum<=99999999"  
  if nowsales<>"" then
  	sql=sql&" and sales='"&nowsales&"'"
  end if
  if nowcustomer<>"" then
  	sql=sql&" and customer='"&nowcustomer&"'"
  end if  
  if nowkeyword<>"" then
  	sql=sql&" and shipmentnum="&nowkeyword
  end if  
  if nowfrom<>"" and nowto="" then
	sql=sql&" and boarddate=#"&nowfrom&"#"
  elseif nowfrom="" and nowto<>"" then
	sql=sql&" and boarddate=#"&nowto&"#"
  elseif nowfrom<>"" and nowto<>"" then
	sql=sql&" and boarddate>=#"&nowfrom&"# and boarddate<=#"&nowto&"#"
  end if
  
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../images/r_1.gif" alt="" /></td>
<td width="100%" background="../images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td width="10%" align="left">&nbsp;船期表查询</td>
<!--	  <td align="right"><input type="button" value="打印客户船期表" class="button" onClick="javascript:var win=window.open('../produit/print_customer_shipment.asp?shipmentnum=10000011','详细信息'); win.focus()"></td>-->
	  <td align="right"><input type="button" value="打印客户船期表" class="button" onClick="print_customer_shipment()"></td>
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
<form name="form1" action="produit_del.asp">
  <input type="hidden" name="queryform" value="<%=request("queryform")%>">
  <input type="hidden" name="form" value="<%=request("form")%>">
  <input type="hidden" name="field" value="<%=request("field")%>">
  <input type="hidden" name="field2" value="<%=request("field2")%>">
  <input type="hidden" name="field3" value="<%=request("field3")%>">

  <thead class="sorting">
	<td class="category" width="100" height="30">船期表号码</td>
	<td class="category setvisible" width="80" height="30">状态</td>
	<td class="category setvisible" width="100" height="30">客户名称</td>
	<td class="category setvisible ininovis" width="100" height="30">供应商名称</td>
	<td class="category setvisible ininovis" width="120" height="30">合同号</td>
	<td class="category" width="80" height="30">项目号</td>	
	<td class="category" width="80" height="30">品名</td>
	<td class="category" width="80" height="30">国家</td>
	<td class="category" width="80" height="30">厂号</td>
	<td class="category" width="100" height="30">动检公司</td>
	<td class="category" width="100" height="30">自动公司</td>
	<td class="category" width="80" height="30">到港期</td>
	<td class="category" width="80" height="30">交单日期</td>
	<td class="category" width="80" height="30">箱号</td>
  <td class="category" width="100" height="30">船公司名</td>
  <td class="category">修改</td>
  <td class="category">选择</td>
  </thead>
  <!--</tr>-->
  <tbody>
  <%
  set rs_shipment =server.createobject("ADODB.RecordSet")	
  rs_shipment.open sql,conn,1,1
  if not rs_shipment.eof then
  do while rs_shipment.eof=false
  %>

    <tr align="center" 
        onDblClick="window.opener.document.<%=request("queryform")%>.refshipment.value='<%=rs_shipment("shipmentnum")%>';
                    window.opener.document.<%=request("queryform")%>.refitem.value='<%=rs_shipment("itemnum")%>';
                    window.opener.document.<%=request("queryform")%>.material.value='<%=rs_shipment("material")%>';
                    window.opener.document.<%=request("queryform")%>.spec.value='<%=rs_shipment("spec")%>';
                    window.opener.document.<%=request("queryform")%>.quantity.value='<%=rs_shipment("casenumber")%>';
                    window.opener.document.<%=request("queryform")%>.weight.value='<%=rs_shipment("contractweight")%>';
                    window.opener.document.<%=request("queryform")%>.guobie.value='<%=rs_shipment("country")%>';
                    window.opener.document.<%=request("queryform")%>.plant.value='<%=rs_shipment("plant")%>';      
                    window.opener.document.<%=request("queryform")%>.boarddate.value='<%=rs_shipment("boarddate")%>';
                    window.opener.document.<%=request("queryform")%>.deliveryport.value='<%=rs_shipment("destination")%>';
                    window.opener.document.<%=request("queryform")%>.package.value='<%=rs_shipment("package")%>';

                    window.close();" 
        <%if rs_shipment("status")="进库" then%>bgcolor="darkgrey"<%elseif rs_shipment("status")="通关中" then%>bgcolor="lawngreen"<%elseif rs_shipment("status")="已送货" then%>bgcolor="red"<%end if%>>     

  <td align="center" height="30"><%=rs_shipment("shipmentnum")%></td>    
  <td align="center" class="setvisible"><%=rs_shipment("status")%></td>	  
  <td align="center" class="setvisible"><%=rs_shipment("customer")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("vendor")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("contract")%></td>
  <td align="center"><%=rs_shipment("itemnum")%></td>  
  <td align="center"><%=rs_shipment("material")%></td>
  <td align="center"><%=rs_shipment("country")%></td>
  <td align="center"><%=rs_shipment("plant")%></td>
  <td align="center"><%=rs_shipment("dongjiancom")%></td>
  <td align="center"><%=rs_shipment("zidongcom")%></td>
  <td align="center"><%=rs_shipment("boarddate")%></td>
  <td align="center"><%=rs_shipment("deliverydate")%></td>
  <td align="center"><%=rs_shipment("case")%></td>
  <td align="center"><%=rs_shipment("carrier")%></td>
  <td align="center">
    	<a href="change_shipment_new.asp?form=<%=request("form")%>&shipment=<%=rs_shipment("shipmentnum")%>&keyword=<%=nowkeyword%>"><img src="../images/res.gif" border="0" hspace="2" align="absmiddle">修改</a>
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
  else
  %>
  <tr align="center" onMouseOver="this.className='highlight'" onMouseOut="this.className=''">
    <td colspan="15" height="25" align="center" style="color:red"><b>没有找到记录</b></td>
  </tr>
  <%
  end if
  %>
  </tbody>
  <%
  if rs_shipment.recordcount>0 then 
  %> 
</table></td></tr>
<%end if%> 
</form>   
</table>
<!--endprint-->
</td>
<td></td>
</tr>

</table>

</body>
</html>