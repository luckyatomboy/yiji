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
<title><%=dianming%> - ��ӡ</title>
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
set rs_shipment=conn.execute("select * from shipment where shipmentnum="&request("shipmentnum"))
%>

<table width="94%" border="0" cellpadding="0" cellspacing="2" align="center" style="margin-top:18pt;margin-left:30pt;">
  <tr> 
    <td height="21" align="center"><img src="../images/print.jpg" align="absmiddle" style="cursor:hand;" onClick="preview();window.close()"></td>
  </tr>
</table>
<!--startprint-->
<table width="94%" border="0" cellpadding="0" cellspacing="2" align="center" style="margin-top:18pt;margin-left:30pt;">
  <tr> 
    <td width="100%" align="center"><h2>��Ʊ</h2></td>
  </tr>
  <tr> 
    <td width="100%" align="center"><h2>INVOICE</h2></td>
  </tr>  
</table>
<table width="94%" border="0" cellpadding="0" cellspacing="2" align="center" style="margin-top:18pt;margin-left:30pt;">
  <tr> 
    <th width="18%" height="30" align="left">�ͻ����ƣ�</th>
	  <td width="54%"><input type="text" value="<%=rs_shipment("customer")%>" size="20" style="border-bottom:solid 0 #000000;border-top:solid 0 #000000;border-left:solid 0 #000000;border-right:solid 0 #000000;"></td>
	  <th width="18%" align="left">��Ʊ��ţ�</th>
	  <td width="10%"><input type="text" value="<%=rs_shipment("shipmentnum")%>" size="20" style="border-bottom:solid 0 #000000;border-top:solid 0 #000000;border-left:solid 0 #000000;border-right:solid 0 #000000;"></td>
	</tr>
  <tr> 
    <th align="left">��ͬ��ţ�</th>
	  <td><input type="text" value="<%=rs_shipment("contract")%>" size="20" style="border-bottom:solid 0 #000000;border-top:solid 0 #000000;border-left:solid 0 #000000;border-right:solid 0 #000000;"></td>
	  <th align="left">���ڣ�</th>
	  <td><input type="text" value="<%=year(now())%>��<%=month(now())%>��<%=day(now())%>��" size="20" style="border-bottom:solid 0 #000000;border-top:solid 0 #000000;border-left:solid 0 #000000;border-right:solid 0 #000000;"></td>
	</tr>	
  <tr> 
    <th align="left">��ţ�</th>
	  <td><input type="text" value="<%=rs_shipment("case")%>" size="20" style="border-bottom:solid 0 #000000;border-top:solid 0 #000000;border-left:solid 0 #000000;border-right:solid 0 #000000;"></td>
	  <th align="left">�����ڣ�</th>
	  <td><input type="text" value="<%=rs_shipment("boarddate")%>" size="20" style="border-bottom:solid 0 #000000;border-top:solid 0 #000000;border-left:solid 0 #000000;border-right:solid 0 #000000;"></td>
	</tr>		
  <tr> 
    <td>&nbsp;</td>
	  <td></td>
	  <td></td>
	  <td></td>
	</tr>		
  <tr> 
    <td></td>
	  <th align="left">�뽫���������</th>
	  <td></td>
	  <td></td>
	</tr>			
	<%
		sql="select * from agent where company='"&rs_shipment("agent")&"'"
		set rs_agent=conn.execute(sql)
		if rs_agent.eof=false then
  %>
  <tr> 
    <td></td>
	  <th align="left"><%=rs_agent("bankname")%></th>
	  <td></td>
	  <td></td>
	</tr>			
  <tr> 
    <td></td>
	  <th align="left">�˺�: <%=rs_agent("bankaccount")%></th>
	  <td></td>
	  <td></td>
	</tr>		
  <tr> 
    <td></td>
	  <th align="left">�˻�: <%=rs_agent("fullname")%></th>
	  <td></td>
	  <td></td>
	</tr>				
<%end if %>
  <tr> 
    <td>&nbsp;</td>
	  <td></td>
	  <td></td>
	  <td></td>
	</tr>		
  <tr> 
    <td>&nbsp;</td>
	  <td></td>
	  <td></td>
	  <td></td>
	</tr>			
</table>
<table width="94%" border="1" cellspacing="0" cellpadding="2" align="center" bordercolor="#000000" style="margin-top:18pt;margin-left:30pt;">
  <tr>
  <th>��Ʒ����</th>
	<th>����</th>
	<th>����</th>
	<th>����</th>
	<th>����</th>
	<th>����</th>
	<th>��λ</th>
  </tr>
  <%
  	if rs_shipment("material")<>"" then
  %>
  <tr>
  	<td align="center"><%=rs_shipment("material")%></td>
  	<td align="center"><%=rs_shipment("memo1")%></td>
  	<td align="center"><%=rs_shipment("country")%></td>
  	<td align="center"><%=rs_shipment("casenumber")%></td>
  	<td align="center"><%=rs_shipment("actualweight1")%></td>
  	<td align="center"><%=rs_shipment("salesprice1")*1000%></td>
  	<td align="center"><%=rs_shipment("trancurrency")%>/��</td>
  </tr>
	<%
		end if
  	if rs_shipment("material2")<>"" then
  %>
  <tr>
  	<td align="center"><%=rs_shipment("material2")%></td>
  	<td align="center"><%=rs_shipment("memo2")%></td>
  	<td align="center"><%=rs_shipment("country")%></td>
  	<td align="center"><%=rs_shipment("casenumber")%></td>
  	<td align="center"><%=rs_shipment("actualweight2")%></td>
  	<td align="center"><%=rs_shipment("salesprice2")*1000%></td>
  	<td align="center"><%=rs_shipment("trancurrency")%>/��</td>
  </tr>
	<%
		end if
  	if rs_shipment("material3")<>"" then
  %>
  <tr>
  	<td align="center"><%=rs_shipment("material3")%></td>
  	<td align="center"><%=rs_shipment("memo3")%></td>
  	<td align="center"><%=rs_shipment("country")%></td>
  	<td align="center"><%=rs_shipment("casenumber")%></td>
  	<td align="center"><%=rs_shipment("actualweight3")%></td>
  	<td align="center"><%=rs_shipment("salesprice3")*1000%></td>
  	<td align="center"><%=rs_shipment("trancurrency")%>/��</td>
  </tr>
	<%
		end if
  	if rs_shipment("material4")<>"" then
  %>
  <tr>
  	<td align="center"><%=rs_shipment("material4")%></td>
  	<td align="center"><%=rs_shipment("memo4")%></td>
  	<td align="center"><%=rs_shipment("country")%></td>
  	<td align="center"><%=rs_shipment("casenumber")%></td>
  	<td align="center"><%=rs_shipment("actualweight4")%></td>
  	<td align="center"><%=rs_shipment("salesprice4")*1000%></td>
  	<td align="center"><%=rs_shipment("trancurrency")%>/��</td>
  </tr>
	<%
		end if
  	if rs_shipment("material5")<>"" then
  %>
  <tr>
  	<td align="center"><%=rs_shipment("material5")%></td>
  	<td align="center"><%=rs_shipment("memo5")%></td>
  	<td align="center"><%=rs_shipment("country")%></td>
  	<td align="center"><%=rs_shipment("casenumber")%></td>
  	<td align="center"><%=rs_shipment("actualweight5")%></td>
  	<td align="center"><%=rs_shipment("salesprice5")*1000%></td>
  	<td align="center"><%=rs_shipment("trancurrency")%>/��</td>
  </tr>
	<%
		end if
  %>								
</table>
<tr>&nbsp;</tr>
<table width="94%" border="0" cellpadding="0" cellspacing="2" align="center" style="margin-top:18pt;margin-left:30pt;">
  <tr> 
    <th width="12%" height="30" align="left">����</th>
	  <th width="17%" align="right"><%=rs_shipment("exchange")%></th>
	  <th width="3%" align="left"></th>
	  <th width="14%" align="left"></th>
	  <th width="17%" align="right"></th>
	  <th width="3%" align="left"></th>
	  <th width="14%" align="left"></th>
	  <th width="17%" align="right"></th>	  
	  <th width="3%" align="left"></th>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
  <tr> 
    <th align="left">���з���</th>
	  <th align="right"><%=rs_shipment("bankfare")%></th>
	  <th align="left"></th>
	  <th align="left">����˾�ļ���</th>
	  <th align="right"><%=rs_shipment("agentfare")%></th>
	  <th align="left"></th>
	  <th align="left">Ӷ��</th>
	  <th align="right"><%=rs_shipment("commission")%></th>	  
	  <th align="left"></th>
	</tr>	
	<tr>
		<td>&nbsp;</td>
	</tr>	
  <tr> 
    <th align="left">�ܼ���</th>
	  <th align="right"><%=rs_shipment("casenumber")%></th>
	  <th align="left"></th>
	  <th align="left">������</th>
	  <th align="right"><%=rs_shipment("netweight")%></th>
	  <th align="left"></th>
	  <th align="left">�ܽ��</th>
	  <th align="right"><%=rs_shipment("actrevenue")%></th>	 
	  <th align="left"></th> 
	</tr>	
</table>
<!--endprint-->
</body>
</html>