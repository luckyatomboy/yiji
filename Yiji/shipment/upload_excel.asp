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
<title><%=dianming%> - 导入船期表</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style/style.css" rel="stylesheet" type="text/css">
<style>
body {
	background-color:#FFFFFF;
}
</style>
</HEAD>

<BODY>
	
<script language="javascript">
function check1(){
if (document.form1.excelfile.value=="")
{
alert("请选择一个文件！");
return false;
}
}
</script>

<%
if request("hid1")="ok" then
dim connExcel

 set connExcel=CreateObject("ADODB.Connection")
 connExcel.Open "Provider=Microsoft.ACE.OLEDB.12.0;Extended properties='Excel 12.0;HDR=Yes; IMEX=1';Data Source="&request("excelfile")
 sql = "select * FROM [Sheet1$]"
 set rs = connExcel.execute(sql)

dim eShipID
dim iShipID, iShipIDA, iShipIDB,iShipItem

iShipID = 0
iShipIDA = 0
iShipIDB = 0
iShipItem = 0

 response.write "正在导入文件"&request("excelfile")&"..."
 response.write "<table border='1'>"
 response.write "<tr>"  
 response.write "<td>我司编号</td>"  
 response.write "<td>成功/失败</td>"   
 response.write "</tr>"  
 
 while not rs.eof
 response.write "<tr>"  
 response.write "<td>" & rs(0) & "</td>" 
 
 if eShipID <> rs(0) then
'如果我司编号变化，重置船期编号和项目编号
 	eShipID = rs(0)   'reset eShipID
 	iShipItem = 1     'reset shipment item num
 
 if rs(1)="A" then
 	 eTranTYpe = "代理"
 	 if iShipIDA = 0 then
 	 	 iShipIDA = 10000001
 	 else
 	 	 iShipIDA = iShipIDA + 1
 	 end if
 	 iShipID = iShipIDA
 else
 	 eTranType = "自营"
 	 if iShipIDB = 0 then
 	 	 iShipIDB = 20000001
 	 else
 	 	 iShipIDB = iShipIDB + 1
 	 end if 	 
 	 iShipID = iShipIDB
 end if
 eContract = rs(2)
 eMaterial = rs(3)

 sqlHead = "insert into shipmentcopy(shipmentnum,trantype,contract) values("&iShipID&",'"&eTranType&"','"&eContract&"')"
 sqlItem = "insert into shipmentitemcopy(shipmentnum,itemnum,material) values("&iShipID&","&iShipItem&",'"&eMaterial&"')"
 conn.execute(sqlHead)
 conn.execute(sqlItem)
 
else
'如果我司编号不变，即为拼柜的项目，不更新主船期表，只插入新项目	
	iShipItem = iShipItem + 1
	
	eMaterial = rs(3)
 
 sqlItem = "insert into shipmentitemcopy(shipmentnum,itemnum,material) values("&iShipID&","&iShipItem&",'"&eMaterial&"')"
 conn.execute(sqlItem)  
 
end if
 
 response.write "<td>成功</td> </tr>" 
 rs.movenext
 wend
 
  response.write "</table>"

set connExcel = Nothing
'end if

else
%>



<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../images/r_1.gif" alt="" /></td>
<td width="100%" background="../images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td>&nbsp;上传Excel文件到数据库 </td>
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
	      <tr>
	        <td width="20%" align="right" height="30">请选择文件：</td>
	        <td width="80%" class="category">
	        	<input name="excelfile" type="file">	
	        </td>
	      </tr>
	      <tr>
			    <td height="30">&nbsp;</td>
		      <td class="category">
					  <input type="submit" value=" 确认上传 " onClick="return check1()" class="button">
					  <input type="hidden" name="hid1" value="ok">
					</td>
	      </tr>     
			</table>
		</form>
</td>
</tr>
</table>

<%
end if
%>

</body>
</html>