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
<title><%=dianming%> - 打印客户船期表</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style2.css" rel="stylesheet" type="text/css">
<script language=javascript>
function preview() { 
bdhtml=window.document.body.innerHTML; 
sprnstr="<!--startprint-->"; 
eprnstr="<!--endprint-->"; 
prnhtml=bdhtml.substr(bdhtml.indexOf(sprnstr)+17); 
prnhtml=prnhtml.substring(0,prnhtml.indexOf(eprnstr)); 
window.document.body.innerHTML=prnhtml; 
window.print(); 
window.document.body.innerHTML=bdhtml; 
         }
</script>
</HEAD>

<BODY>
<%
  shipmentid=split(request("shipmentnum"),",")
  for i=0 to UBound(shipmentid)
    response.write(shipmentid(i))
    shipmentnum=left(shipmentid(i),8)
    itemnum=right(shipmentid(i),1)
    if clause="" then
      clause = "( b.shipmentnum="&shipmentnum&" and b.itemnum="&itemnum&")"
    else
      clause = clause & " or ( b.shipmentnum="&shipmentnum&" and b.itemnum="&itemnum&")"
    end if
  next  
  response.write(clause)
set rs_shipment=conn.execute("select a.shipmentnum, a.vendor, a.country, a.plant, b.itemnum, b.customer, b.material, b.casenumber, b.actualnetweight, b.purchaseprice, b.invoicecurrency from shipment a inner join shipmentitem b on a.shipmentnum = b.shipmentnum where "&clause)
  'a.shipmentnum="&request("shipmentnum"))
%>


<table width="94%" border="0" cellpadding="0" cellspacing="2" align="center" style="margin-top:18pt;margin-left:30pt;">
  <tr> 
    <td height="21" align="center"><img src="../images/print.jpg" align="absmiddle" style="cursor:hand;" onClick="preview();window.close()"></td>
  </tr>
</table>
<!--startprint-->
<table width="94%" border="0" cellpadding="0" cellspacing="2" align="center" style="margin-top:18pt;margin-left:30pt;">
  <tr> 
    <td width="100%" align="center"><h2>客户船期表</h2></td>
  </tr>
</table>
<table width="94%" border="0" cellpadding="0" cellspacing="2" align="center" style="margin-top:18pt;margin-left:30pt;">
  <tr> 
    <th width="18%" height="30" align="left">客户名称：</th>
	  <td width="54%"><input type="text" value="<%=rs_shipment("customer")%>" size="20" style="border-bottom:solid 0 #000000;border-top:solid 0 #000000;border-left:solid 0 #000000;border-right:solid 0 #000000;"></td>
	</tr>			
</table>
<table width="94%" border="1" cellspacing="0" cellpadding="2" align="center" bordercolor="#000000" style="margin-top:18pt;margin-left:30pt;">
  <tr>
  <th>我司编号</th>
  <th>供应商</th>
  <th>国家</th>
	<th>厂号</th>
	<th>中文品名</th>
	<th>箱数</th>
	<th>重量</th>
	<th>单价</th>
	<th>单位</th>
  </tr>
  <tr>
  	<td align="center"><%=rs_shipment("shipmentnum")%></td>
  	<td align="center"><%=rs_shipment("vendor")%></td>
  	<td align="center"><%=rs_shipment("country")%></td>
  	<td align="center"><%=rs_shipment("plant")%></td>
  	<td align="center"><%=rs_shipment("material")%></td>
  	<td align="center"><%=rs_shipment("casenumber")%></td>
  	<td align="center"><%=rs_shipment("actualnetweight")%></td>
  	<td align="center"><%=rs_shipment("purchaseprice")*1000%></td>
  	<td align="center"><%=rs_shipment("invoicecurrency")%>/吨</td>
  </tr>						
</table>
</table>
<!--endprint-->
</body>
</html>