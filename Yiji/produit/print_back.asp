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
<title><%=dianming%> - 打印</title>
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
set rs_buy=conn.execute("select * from buy where zu=false and bianhao='"&request("bianhao")&"'")
nowid_login=rs_buy("id_login")
set rs_login=conn.execute("select * from login where id="&rs_buy("id_login"))
set rs_huiyuan=conn.execute("select * from huiyuan where id="&rs_buy("id_huiyuan"))
%>

<table width="98%" border="0" cellpadding="0" cellspacing="2" align="center">
  <tr> 
    <td height="21" align="center"><img src="../images/print.jpg" align="absmiddle" style="cursor:hand;" onClick="preview();window.close()"></td>
  </tr>
</table>
<!--startprint-->
<table width="98%" border="0" cellpadding="0" cellspacing="2" align="center">
  <tr> 
    <td height="21" align="center"><h1>客户退货单</h1></td>
  </tr>
  <tr> 
    <td height="21" align="right">
	  客户名称：<input type="text" value="<%if rs_huiyuan.eof=false then%><%=rs_huiyuan("username")%><%end if%>" size="20" style="border-bottom:solid 1 #000000;border-top:solid 0 #000000;border-left:solid 0 #000000;border-right:solid 0 #000000;">&nbsp;&nbsp;
	  退货日期：<input type="text" value="<%=rs_buy("selldate")%>" size="20" style="border-bottom:solid 1 #000000;border-top:solid 0 #000000;border-left:solid 0 #000000;border-right:solid 0 #000000;">&nbsp;&nbsp;
	  发票号码：<input type="text" value="<%=request("bianhao")%>" size="20" style="border-bottom:solid 1 #000000;border-top:solid 0 #000000;border-left:solid 0 #000000;border-right:solid 0 #000000;">&nbsp;&nbsp;
	</td>
  </tr>
</table>
<table width="98%" border="1" cellspacing="0" cellpadding="2" align="center" bordercolor="#000000">
  <tr>
	<%if showpic="yes" then%>
	<th width="70">图片</th>
	<%end if%>
    <th height="30">货号</th>
	<th>产品名称</th>
	<th>数量</th>
	<th>单位</th>
	<th>单价</th>
	<th>金额</th>
  </tr>
  <%
  totalprice=0
  do while rs_buy.eof=false
  set rs_produit=conn.execute("select * from produit where huohao='"&rs_buy("huohao")&"'") 
  %>
  <tr>
    <%if showpic="yes" then%><td align="center"><%if rs_produit.eof then%><%if rs_buy("photo")<>"" then%><a href="../upload/<%=rs_buy("photo")%>" target="_blank"><img src="../upload/<%=rs_buy("photo")%>" border="0" width="60"></a><%else%>无图<%end if%><%else%><%if rs_produit("photo")<>"" then%><a href="../upload/<%=rs_produit("photo")%>" target="_blank"><img src="../upload/<%=rs_produit("photo")%>" border="0" width="60"></a><%else%>无图<%end if%><%end if%></td><%end if%>
    <td align="center" height="25"><%if rs_produit.eof then%><%=rs_buy("huohao")%><%else%><%=rs_produit("huohao")%><%end if%></td>
	<td align="center"><%if rs_produit.eof then%><%=rs_buy("title")%><%else%><%=rs_produit("title")%><%end if%></td>
	<td align="center"><%=rs_buy("shulian")%></td>
	<td align="center"><%if rs_produit.eof=false then%><%=rs_produit("danwei")%><%end if%></td>
	<td align="center"><%=formatnumber(rs_buy("price"),3)%></td>
	<td align="center"><%=formatnumber(rs_buy("price")*rs_buy("shulian"),3)%></td>
  </tr>
  <%
  totalprice=totalprice+rs_buy("price")*rs_buy("shulian")
  rs_buy.movenext
  loop
  %>
  <tr>
    <td colspan="<%if showpic="yes" then%>6<%else%>5<%end if%>" align="right">总计 :&nbsp;</td>
	<td align="center"><%=formatnumber(totalprice,3)%></td>
  </tr>
</table>
<table width="98%" border="0" cellpadding="0" cellspacing="2" align="center">
  <tr> 
    <td height="21" width="10%">
	  &nbsp;;
	</td>
    <td align="right" width="90%">
	  经办人：<input type="text" value="<%if rs_login.eof then%><%=rs_buy("login")%><%else%><%=rs_login("username")%> (<%=rs_login("bianhao")%>)<%end if%>" size="20" style="border-bottom:solid 1 #000000;border-top:solid 0 #000000;border-left:solid 0 #000000;border-right:solid 0 #000000;">&nbsp;&nbsp;
	  收款人：<input type="text" value="<%if rs_login.eof then%><%=rs_buy("login")%><%else%><%=rs_login("username")%> (<%=rs_login("bianhao")%>)<%end if%>" size="20" style="border-bottom:solid 1 #000000;border-top:solid 0 #000000;border-left:solid 0 #000000;border-right:solid 0 #000000;">&nbsp;&nbsp;
	</td>
  </tr>
</table>
<!--endprint-->
</body>
</html>