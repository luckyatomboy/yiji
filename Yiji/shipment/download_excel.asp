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
<title><%=dianming%> - �������ڱ�</title>
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
alert("��ѡ��һ���ļ���");
return false;
}
}
</script>

<%
if request("hid1")="ok" then
dim s,sql,filename,fs,myfile,x 
  
Set fs = server.CreateObject("scripting.filesystemobject") 
'--�������������ɵ�EXCEL�ļ������µĴ�� 
filename = request("excelfile")
'--���ԭ����EXCEL�ļ����ڵĻ�ɾ���� 
if fs.FileExists(filename) then 
   fs.DeleteFile(filename) 
end  if 
'--����EXCEL�ļ� 
set myfile = fs.CreateTextFile(filename,true) 

sql = "select * from shipmentcopy"
Set rs = conn.execute(sql)

sqlItem = "select top 1 * from shipmentitemcopy "
Set rsitem = conn.execute(sqlItem)   

 response.write "���ڵ����ļ�"&request("excelfile")&"..."
 response.write "<table border='1'>"
 response.write "<tr>"  
 response.write "<td>���ڱ��</td>"  
 response.write "<td>��Ŀ���</td>"  
 response.write "<td>�ɹ�/ʧ��</td>"   
 response.write "</tr>"  

if not rs.EOF and not rs.BOF then 
  
   dim  trLine,responsestr,index,iShipNum,iShipItem
   strLine=""
   index = 1
'�����ڱ��ֶ���
   For each x in rs.fields 
     strLine = strLine & x.name & chr(9) 
   Next
'������ϸ���ֶ������ų����ڱ��
   For each x in rsitem.fields 
   	 if index > 1 then
     	strLine = strLine & x.name & chr(9) 
     end if
     index = index + 1
   Next   
  
'--�����������д��EXCEL 
   myfile.writeline strLine 
  
   Do while Not rs.EOF 
     strLine=""
     headLine=""
     index = 1
  
     for each x in rs.Fields 
       strLine = strLine & x.value &  chr(9) 
     	 if index = 1 then 
     	 	 iShipNum = x.value
     	 end if
     	 index = index + 1       
     next 
     
     headLine = strLine
     
		sqlItem = "select * from shipmentitemcopy where shipmentnum = "&rs(0)
		Set rsitem = conn.execute(sqlItem)     
		
		do while not rsitem.eof
			strLine=""
			
		 index = 1
     
     for each x in rsitem.Fields 
     	 if index > 1 then
         strLine = strLine & x.value &  chr(9) 
       end if
     	 if index = 2 then 
     	 	 iShipItem = x.value
     	 end if
     	 index = index + 1       
     next 
     strLine = headLine & strLine
          
     myfile.writeline  strLine 
     
		 response.write "<tr>"  
		 response.write "<td>"&iShipNum&"</td>"  
		 response.write "<td>"&iShipItem&"</td>"   
		 response.write "<td>�ɹ�</td>"   
		 response.write "</tr>"  
		 
		 rsitem.movenext
 
 		loop
   
     rs.MoveNext 
   loop 
  
end if 
 response.write "</table>"  

else
%>



<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../images/r_1.gif" alt="" /></td>
<td width="100%" background="../images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td>&nbsp;�������ڱ�ΪExcel�ļ� </td>
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
	        <td width="20%" align="right" height="30">�������ļ�����·����</td>
	        <td width="80%" class="category">
	        	<input name="excelfile">	
	        </td>
	      </tr>
	      <tr>
			    <td height="30">&nbsp;</td>
		      <td class="category">
					  <input type="submit" value=" ȷ�ϵ��� " onClick="return check1()" class="button">
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