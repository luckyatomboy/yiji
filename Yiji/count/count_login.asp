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
<title><%=dianming%> - Ա������ͳ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style/style.css" rel="stylesheet" type="text/css">
<style>
body {
	background-color:#FFFFFF;
}
</style>
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
if fla34="0" and session("redboy_id")<>"1" then
%>
<br><center><img src="../images/note.gif" align="absmiddle">&nbsp;<font color="#FF0000">�㲻�߱���Ȩ�ޣ��������Ա��ϵ��</font></center>
<%  
  response.end
end if

'ȡ�������ؼ���  
nowstartdate=request("startdate") 
if nowstartdate="" then
  nowstartdate=date()-30
end if
nowenddate=request("enddate") 
if nowenddate="" then
  nowenddate=date()
end if
nowkeyword=request("keyword") 
%>
<table width="100%" border="0" cellpadding="0" cellspacing="2" align="center">
<form name="form2">
  <tr> 
    <td width="5%" height="21">&nbsp;<img src="../images/print.jpg" align="absmiddle" style="cursor:hand;" onClick="preview();window.close()"></td>
	<td width="95%" align="right">
	  ��ʼ���ڣ�
      <input name="startdate" value="<%=nowstartdate%>" readonly style="width:100px">
	  <img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('day.asp?form=form2&field=startdate&oldDate='+startdate.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">
	  �������ڣ�
      <input name="enddate" value="<%=nowenddate%>" readonly style="width:100px">
	  <img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('day.asp?form=form2&field=enddate&oldDate='+enddate.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=740');">
	  <input type="text" name="keyword" size="20" value="<%=nowkeyword%>">
	  <input type="hidden" name="hid" value="ok">
	  <input type="submit" value=" ��ѯ " class="button">
	  &nbsp;&nbsp;
	</td>
  </tr>
</form>  
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../images/r_1.gif" alt="" /></td>
<td width="100%" background="../images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td>&nbsp;Ա������ͳ��<a href="excel_login.asp?page=<%=currentpage%>&startdate=<%=nowstartdate%>&enddate=<%=nowenddate%>&ku=<%=nowku%>&keyword=<%=nowkeyword%>&order1=<%=request("order1")%>&order2=<%=request("order2")%>&order3=<%=request("order3")%>&order4=<%=request("order4")%>&order5=<%=request("order5")%>&order6=<%=request("order6")%>&order7=<%=request("order7")%>&order8=<%=request("order8")%>&order9=<%=request("order9")%>&order10=<%=request("order10")%>&order11=<%=request("order11")%>&order12=<%=request("order12")%>&order13=<%=request("order13")%>&order14=<%=request("order14")%>&order15=<%=request("order15")%>" target="_blank"><img src="../images/excel.jpg" border="0" align="absmiddle" alt="����Excel���"></a></td>
	  <td align="right">&nbsp;</td>

    </tr>
  </table>
</td>
<td><img src="../images/r_2.gif" alt="" /></td>
</tr>
<tr>
<td></td>
<td>
<!--startprint-->
<table align="center" cellpadding="4" cellspacing="1" class="toptable grid" border="1">
  <tr align="center">
    <td class="category" height="30">Ա�����</td>
	<td class="category">Ա������</td>
    <td class="category">��������</td>
	<td class="category">�������</td>
    <td class="category">�ϼƹ���</td>
    <td class="category">&nbsp;</td>
  </tr>
  <%
  sql="select * from login where 1=1"
  if nowkeyword<>"" then
    sql=sql&" and (username like '%"&nowkeyword&"%' or bianhao like '%"&nowkeyword&"%')"
  end if  
  sql=sql&" order by id"
  set rs_login=conn.execute(sql)
  do while rs_login.eof=false
  %>
  <tr onMouseOver="this.className='highlight'" onMouseOut="this.className=''">
    <td align="center" height="25"><%=rs_login("bianhao")%></td>
	<td align="center"><%=rs_login("username")%></td>
    <td align="center"><%=rs_login("gongzi")/30*abs(cdate(nowenddate)-cdate(nowstartdate))%> Ԫ</td>
    <td align="center">
		  <%
		    sql="select * from sell where zu=false and isok and id_login="&rs_login("id")&" and selldate>=#"&nowstartdate&"# and selldate<=#"&nowenddate&"#"
			set rs_sell=conn.execute(sql)
			nowtichen=0
			do while rs_sell.eof=false
			  sql="select * from produit where id="&rs_sell("id_produit")
			  set rs_produit=conn.execute(sql)
			  if rs_produit.eof=false then
			    if rs_produit("tichen_type")=0 then
				  nowtichen=nowtichen+(rs_produit("tichen")*rs_sell("price")*rs_sell("shulian")/100)
				else
			      nowtichen=nowtichen+rs_produit("tichen")*rs_sell("shulian")
				end if
			  end if
			  rs_sell.movenext
			loop		  		
		  %>
		  <%=formatnumber(nowtichen,3)%>&nbsp;Ԫ 	
	</td>
    <td align="center"><%=formatnumber(rs_login("gongzi")+nowtichen,3)%>&nbsp;Ԫ</td>	
    <td align="center" width="120"><input type="button" value="��ϸ���ۼ�¼" onClick="javascript:var win=window.open('user_show.asp?id=<%=rs_login("id")%>','Ա����ϸ��Ϣ','width=853,height=470,top=176,left=161,toolbar=no,location=no,directories=no,status=no,menubar=no,resizable=no,scrollbars=yes'); win.focus()" class="button"></td>
  </tr>
  <%
    rs_login.movenext
  loop
  %>
</table>
<!--endprint-->
</td>
<td></td>
</tr>
<tr>
<td><img src="../images/r_4.gif" alt="" /></td>
<td></td>
<td><img src="../images/r_3.gif" alt="" /></td>
</tr>
</table>
</body>
</html>