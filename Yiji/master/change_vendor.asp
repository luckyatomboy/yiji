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
<title><%=dianming%> - �޸Ĺ�Ӧ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style/style.css" rel="stylesheet" type="text/css">
<style>
body {
	background-color:#FFFFFF;
}
</style>
</HEAD>

<BODY>
<%
if fla35="0" and session("redboy_id")<>"1" then
%>
<br><center><img src="../images/note.gif" align="absmiddle">&nbsp;<font color="#FF0000">�㲻�߱���Ȩ�ޣ��������Ա��ϵ��</font></center>
<%  
  response.end
end if
%>

<%if request("hid1")="" then


  sql="select * from locktable where tablename='vendor' and combinedkey='"&request("vendorname")&"'"
  set rs_lock=conn.execute(sql)
  if rs_lock.eof = false then
%>
    <script language="javascript">
    alert("�û�<%=rs_lock("username")%>���ڱ༭�ü�¼�����Ժ����ԣ�");
    window.location.href="master.asp";
    </script> 
<%else
    sql="insert into locktable(tablename,combinedkey,status,username,locktime) values('vendor','"&request("vendorname")&"','E','"&session("redboy_username")&"',#"&now()&"#)"  
    conn.execute(sql)
end if


sql="select * from vendor where vendorname='"&request("vendorname")&"'"
set rs=conn.execute(sql)
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../images/r_1.gif" alt="" /></td>
<td width="100%" background="../images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td>&nbsp;�޸Ĺ�Ӧ������</td>
	  <td align="right">&nbsp;</td>
    </tr>
  </table>
</td>
<td><img src="../images/r_2.gif" alt="" /></td>
</tr>
<tr>
<td></td>
<td>
<table align="center" cellpadding="4" cellspacing="1" class="toptable grid" border="1">
  <form name="form2">
  		<input type="hidden" name="keyword" value="<%=request("keyword")%>">
  		<input type="hidden" name="vendorname" value="<%=request("vendorname")%>">
      <tr>
        <td width="25%" height="30" align="right">��Ӧ�����ƣ�</td>
        <td width="75%" class="category"><%=rs("vendorname")%>
        	</td>
      </tr>    
      <tr>
        <td align="right" height="30">����</td>
        <td class="category">
          <%
          sql="select * from country"
          set rs_country=conn.execute(sql)
          %>
          <select name="country">
            <%
              do while rs_country.eof=false
            %>
              <option value="<%=rs_country("country")%>" <%if rs_country("country")=rs("country") then%>selected="selected"<%end if%>><%=rs_country("country")%></option>
            <%
              rs_country.movenext
              loop
            %>
          </select>
        </td>
      </tr>        
      <tr>
        <td align="right" height="30">��˾��ַ��</td>
        <td class="category"><input type="text" name="address" style="width:300px" value="<%=rs("address")%>"></td>
      </tr> 	  
      <tr>
        <td align="right" height="30">��ϵ�绰��</td>
        <td class="category"><input type="text" name="tel" style="width:200px" value="<%=rs("tel")%>"></td>
      </tr>	 
      <tr>
        <td align="right" height="30">���棺</td>
        <td class="category"><input type="text" name="fax" style="width:200px" value="<%=rs("fax")%>"></td>
      </tr>	       
	    <tr>
        <td align="right" height="30">Email��</td>
        <td class="category"><input type="text" name="email" style="width:200px" value="<%=rs("email")%>"></td>
      </tr>    
      <tr>
        <td align="right" height="30">���ʽ1��</td>
        <td class="category"><input type="text" name="term1" style="width:100px" value="<%=rs("term1")%>"></td>
      </tr> 	
      <tr>
        <td align="right" height="30">���ʽ2��</td>
        <td class="category"><input type="text" name="term2" style="width:100px" value="<%=rs("term2")%>"></td>
      </tr> 	    
<!--added on 2017/2/4  begin-->
      <tr>
        <td align="right" height="30">����1��</td>
        <td class="category"><input type="text" name="plant1" style="width:60px" value="<%=rs("plant1")%>"> 
        	&nbsp;&nbsp;&nbsp;����2 <input type="text" name="plant2" style="width:60px" value="<%=rs("plant2")%>"> 
        	&nbsp;&nbsp;&nbsp;����3 <input type="text" name="plant3" style="width:60px" value="<%=rs("plant3")%>"> 
        	&nbsp;&nbsp;&nbsp;����4 <input type="text" name="plant4" style="width:60px" value="<%=rs("plant4")%>"> 
        </td>
      </tr> 	      
<!--end -->
      <tr>
        <td align="right" height="30">��ע��</td>
        <td class="category"><textarea name="beizhu" cols="70" rows="4"><%=rs("memo")%></textarea></td>
      </tr>	   
      <tr>
        <td align="right" height="30">����ʱ�䣺</td>
        <td class="category"><%=rs("createdate")%>
      </tr>	  
      <tr>
        <td align="right" height="30">�����ˣ�</td>
        <td class="category"><%=rs("creator")%>
      </tr>	   
      <tr>
        <td align="right" height="30">�޸�ʱ�䣺</td>
        <td class="category"><%=rs("changedate")%>
      </tr>	  
      <tr>
        <td align="right" height="30">�޸��ˣ�</td>
        <td class="category"><%=rs("changer")%>
      </tr>	            
      <tr>
	    <td height="30">&nbsp;</td>
        <td class="category">
		  <input type="submit" value=" ȷ���޸� " class="button">&nbsp;&nbsp;&nbsp;&nbsp;
		  <input type="hidden" name="hid1" value="ok">
      <input type="hidden" name="hid2" value="vendor">
			<input type="button" value=" �����޸ķ��� " onClick="if (confirm('ȷ��Ҫ�����޸���')) {window.open('delete_lock_table.asp?tablename=vendor&combinedkey=<%=request("vendorname")%>'); window.location.href='master.asp';}" class="button">&nbsp;&nbsp;&nbsp;&nbsp;
			<%
			if fla35="0" and session("redboy_id")<>"1" then
			else
			%>			
			<input type="button" value=" ɾ�� " onClick="if (confirm('ȷ��Ҫɾ���ù�Ӧ����')) {window.open('delete_vendor.asp?vendorname=<%=request("vendorname")%>'); window.location.href='master.asp';}" class="button"></td>
			<%end if%>				
      </tr>	    
</table>
</td>
<td></td>
</tr>
<tr>
<td><img src="../images/r_4.gif" alt="" /></td>
<td></td>
<td><img src="../images/r_3.gif" alt="" /></td>
</tr>
</table>
</form>
<%
else
nowvendorname=request("vendorname")
nowaddress=request("address")
nowtel=request("tel")
nowfax=request("fax")
nowemail=request("email")
'nowpartner=request("partner")
nowplant1=request("plant1")
nowplant2=request("plant2")
nowplant3=request("plant3")
nowplant4=request("plant4")
nowdes=request("beizhu")
nowkeyword=request("keyword")
nowcountry=request("country")
nowterm1=request("term1")
nowterm2=request("term2")

'��鹤���Ƿ��ظ�'
for i = 1 to 4
  if request("plant"&i) <> "" then
    sql="select * from plant where plantid = '"&request("plant"&i)&"' and country = '"&nowcountry&"' and vendor <> '"&nowvendorname&"'"
    set rs=conn.execute(sql)
    if rs.eof=false then
    %>
    <script language="javascript">
    alert("<%=nowcountry%>����<%=rs("plantid")%>�ѱ���Ӧ��<%=rs("vendor")%>ռ�ã�");
    window.history.go(-1)
    </script> 
    <%
      response.end
    end if
    rs.close
  end if
next 

set rs=server.createobject("ADODB.RecordSet")
sql="select * from vendor where vendorname='"&request("vendorname")&"'"
rs.open sql,conn,1,3

'��ɾ���ɹ�������
for i = 1 to 4
  if rs("plant"&i) <> "" then  
    sql="delete from plant where plantid = '"&rs("plant"&i)&"' and country='"&rs("country")&"'"
    conn.execute(sql)
  end if
next

rs("address")=nowaddress
rs("tel")=nowtel
rs("fax")=nowfax
rs("email")=nowemail
rs("plant1")=nowplant1
rs("plant2")=nowplant2
rs("plant3")=nowplant3
rs("plant4")=nowplant4
rs("country")=nowcountry
rs("term1")=nowterm1
rs("term2")=nowterm2
rs("memo")=nowdes
rs("changedate")=now
rs("changer")=session("redboy_username")
rs.update
rs.close

'�ٲ����¹�������
for i = 1 to 4
  if request("plant"&i) <> "" then  
    sql="insert into plant(plantid,Country,vendor,CreateDate,Creator) values('"&request("plant"&i)&"','"&nowcountry&"','"&nowvendorname&"',#"&now()&"#,'"&session("redboy_username")&"')"
    conn.execute(sql)
  end if
next

sql="delete from locktable where tablename='vendor' and combinedkey='"&request("vendorname")&"'"
conn.execute(sql)

%>
<script language="javascript">
alert("�ͻ������޸ĳɹ���")
window.location.href="query_vendor.asp?form=<%=request("form")%>&keyword=<%=nowkeyword%>"
</script> 
<%
end if
%>
</body>
</html>