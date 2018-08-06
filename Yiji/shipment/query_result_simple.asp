<!-- #include file="../conn2.asp" -->
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
<!--Header-->
	<td class="category" width="100" height="30">船期表号码</td>
	<td class="category setvisible" width="80" height="30">状态</td>
	<td class="category setvisible" width="100" height="30">客户名称</td>
	<td class="category setvisible ininovis" width="100" height="30">供应商名称</td>
	<td class="category setvisible ininovis" width="120" height="30">合同号</td>
	<td class="category setvisible" width="80" height="30">项目号</td>	
	<td class="category setvisible" width="80" height="30">品名</td>
	<td class="category setvisible" width="80" height="30">国家</td>
	<td class="category setvisible" width="80" height="30">厂号</td>

	<td class="category setvisible" width="80" height="30">到港期</td>
	<td class="category setvisible" width="80" height="30">交单日期</td>
	<td class="category setvisible" width="80" height="30">箱号</td>
  <td class="category setvisible" width="100" height="30">船公司名</td>
  <td class="category setvisible ininovis" width="100" height="30">结算单状态</td>
  <td class="category setvisible ininovis" width="80" height="30">业务类型</td>
  <td class="category setvisible ininovis" width="100" height="30">代理公司</td>
  <td class="category setvisible ininovis" width="50" height="30">采购</td>
  <td class="category setvisible ininovis" width="50" height="30">销售</td>
  <td class="category setvisible ininovis" width="50" height="30">跟单</td>
  <td class="category setvisible ininovis" width="50" height="30">品类</td>
  <td class="category setvisible ininovis" width="60" height="30">付款条件</td>    

  <td class="category setvisible ininovis" width="100" height="30">港口</td>
  <td class="category setvisible ininovis" width="100" height="30">靠泊码头</td>
  <td class="category setvisible ininovis" width="100" height="30">船名航次</td>

  <td class="category setvisible ininovis" width="100" height="30">进口批文</td>
  <td class="category setvisible ininovis" width="100" height="30">两证齐备日</td>
  <td class="category setvisible ininovis" width="100" height="30">动检证</td>
  <td class="category setvisible ininovis" width="100" height="30">自动证</td>
  <td class="category setvisible ininovis" width="100" height="30">自动证上报日</td>
  <td class="category setvisible ininovis" width="100" height="30">自动证交批日</td>
  
  <td class="category setvisible ininovis" width="100" height="30">预保日</td>
  <td class="category setvisible ininovis" width="100" height="30">补保日</td>
  <td class="category setvisible ininovis" width="100" height="30">保费支付日</td>     
  <td class="category setvisible ininovis" width="100" height="30">保单</td>    

  <td class="category setvisible ininovis" width="100" height="30">预计装船月份</td>
  <td class="category setvisible ininovis" width="100" height="30">实际装船期</td>
  
  <td class="category setvisible ininovis" width="100" height="30">客户交货期</td>    


  <td class="category setvisible ininovis" width="100" height="30">提单号</td>
  <td class="category setvisible ininovis" width="100" height="30">卫生证号</td>     
  <td class="category setvisible ininovis" width="100" height="30">铅封号</td>     
  <td class="category setvisible ininovis" width="60" height="30">电子信息</td>
  <td class="category setvisible ininovis" width="60" height="30">清真证明</td>     
  <td class="category setvisible ininovis" width="50" height="30">标签</td>     

  <td class="category setvisible ininovis" width="100" height="30">预付款日期</td>
  <td class="category setvisible ininovis" width="100" height="30">预付款金额</td>
  <td class="category setvisible ininovis" width="100" height="30">币种</td>     
  <td class="category setvisible ininovis" width="100" height="30">订金收到日期</td>     
  <td class="category setvisible ininovis" width="60" height="30">成交条款</td>  

  <td class="category setvisible ininovis" width="100" height="30">尾款支付日期</td>
  <td class="category setvisible ininovis" width="100" height="30">免箱期</td>
  <td class="category setvisible ininovis" width="100" height="30">兽医官</td>     
  <td class="category setvisible ininovis" width="100" height="30">放行日期</td>     

  <td class="category setvisible ininovis" width="100" height="30">到单日期</td>
  <td class="category setvisible ininovis" width="100" height="30">预计赎单日期</td>
  <td class="category setvisible ininovis" width="100" height="30">我司赎单日期</td>     
  <td class="category setvisible ininovis" width="100" height="30">代理赎单日期</td>      

  <td class="category setvisible ininovis" width="100" height="30">送货日期</td>
  <td class="category setvisible ininovis" width="100" height="30">送货方向</td>
  

  <td class="category setvisible ininovis" width="100" height="30">保证金状态</td>
  <td class="category setvisible ininovis" width="100" height="30">保证金金额</td>
  <td class="category setvisible ininovis" width="100" height="30">退保证金金额</td>     
  <td class="category setvisible ininovis" width="60" height="30">索赔状态</td>     
  <td class="category setvisible ininovis" width="60" height="30">索赔负责人</td>    

<!--Item-->
  <td class="category setvisible ininovis" width="100" height="30">规格</td>
  <td class="category setvisible ininovis" width="100" height="30">包装</td>
  <td class="category setvisible ininovis" width="100" height="30">合同重量</td>     
  <td class="category setvisible ininovis" width="60" height="30">重量单位</td>     
  <td class="category setvisible ininovis" width="100" height="30">实际净重</td>    
  <td class="category setvisible ininovis" width="100" height="30">采购价格</td>
  <td class="category setvisible ininovis" width="60" height="30">价格单位</td>
  <td class="category setvisible ininovis" width="100" height="30">生产日期</td>     
  <td class="category setvisible ininovis" width="100" height="30">箱数</td>     
  <td class="category setvisible ininovis" width="100" height="30">发票总金额</td>   
  <td class="category setvisible ininovis" width="60" height="30">币种</td>     
  <td class="category setvisible ininovis" width="100" height="30">尾款金额</td> 


  <td class="category">修改</td>
  <td class="category">选择</td>
  </thead>
  <!--</tr>-->
  <tbody>

  <%
  'if nowfieldno<>"" then
  sql=request("sql")
  if sql<>"" then
  set rs_shipment =server.createobject("ADODB.RecordSet") 
  rs_shipment.open sql,conn,1,1
  if not rs_shipment.eof then
  do while rs_shipment.eof=false
  %>

    <tr>     

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
  <td align="center" class="setvisible ininovis"><%=rs_shipment("twodocumentready")%></td>  
  <td align="center" class="setvisible ininovis"><%=rs_shipment("dongjian")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("zidong")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("zidongapplydate")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("zidongreportdate")%></td>

  <td align="center" class="setvisible ininovis"><%=rs_shipment("planinsurance")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("supinsurance")%></td>    
  <td align="center" class="setvisible ininovis"><%=rs_shipment("insurancepayment")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("insurancenumber")%></td>    

  <td align="center" class="setvisible ininovis"><%=rs_shipment("planship")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("shipdate")%></td>    

  <td align="center" class="setvisible ininovis"><%=rs_shipment("customerdeliverydate")%></td>   


  <td align="center" class="setvisible ininovis"><%=rs_shipment("ladnumber")%></td>    
  <td align="center" class="setvisible ininovis"><%=rs_shipment("weishengzheng")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("locknumber")%></td>   
  <td align="center" class="setvisible ininovis"><%if rs_shipment("einformation")="True" then%>有<%else%>无<%end if%></td>
  <td align="center" class="setvisible ininovis"><%if rs_shipment("MuslimCertification")="True" then%>有<%else%>无<%end if%></td>    
  <td align="center" class="setvisible ininovis"><%if rs_shipment("elabel")="True" then%>有<%else%>无<%end if%></td>

  <td align="center" class="setvisible ininovis"><%=rs_shipment("prepaydate")%></td>  
  <td align="center" class="setvisible ininovis"><%=rs_shipment("prepayment")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("trancurrency")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("downpaymentreceiptdate")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("tradingterm")%></td>  

  <td align="center" class="setvisible ininovis"><%=rs_shipment("finalpaymentdate")%></td>  
  <td align="center" class="setvisible ininovis"><%=rs_shipment("freestayperiod")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("shouyi")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("passdate")%></td>  

  <td align="center" class="setvisible ininovis"><%=rs_shipment("documentarrival")%></td>  
  <td align="center" class="setvisible ininovis"><%=rs_shipment("planretirebill")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("internalretirebill")%></td>
  <td align="center" class="setvisible ininovis"><%=rs_shipment("externalretirebill")%></td>    

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
      <a href="change_shipment_new.asp?form=<%=request("form")%>&shipment=<%=rs_shipment("shipmentnum")%>&keyword=<%=nowkeyword%>"><img src="../images/res.gif" border="0" hspace="2" align="absmiddle">修改</a>
  </td>
  <td align="center">
    <input type="checkbox" name="sel" value="<%=rs_shipment("shipmentnum")&rs_shipment("itemnum")%>" style="border:0">
  </td>
<!--  <td align="center"><input type="checkbox" name='chkArr'  id='chkArr' />";  -->
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
  end if
  %>

  </tbody>

</form>   
</table>
<!--endprint-->
</td>
<td></td>
</tr>

</table>