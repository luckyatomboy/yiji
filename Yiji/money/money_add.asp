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
<title><%=dianming%> - �������</title>
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
if fla70="0" and session("redboy_id")<>"1" then
%>
<br><center><img src="../images/note.gif" align="absmiddle">&nbsp;<font color="#FF0000">�㲻�߱���Ȩ�ޣ��������Ա��ϵ��</font></center>
<%  
  response.end
end if
%>
<%
if request("hid2")<>"" then
nowtype=request("type")
nowbigclass=request("bigclass")
nowsmallclass=request("smallclass")
nowprice=request("price")
if nowprice="" then
  nowprice=0
end if
nowbeizhu=request("beizhu")
nowid_login=request("id_login")
set rs_login=conn.execute("select * from login where id="&nowid_login)
sql="insert into redboy_money(id_bigclass,id_smallclass,type,price,beizhu,id_login,login,selldate) values("&nowbigclass&","&nowsmallclass&","&nowtype&","&nowprice&",'"&nowbeizhu&"',"&nowid_login&",'"&rs_login("username")&"',#"&date()&"#)"
conn.execute(sql)
%>
<script language="javascript">
alert("�����ɹ���")
</script> 
<%
end if
%>
<script language="javascript">
function isNumberString (InString,RefString)
{
if(InString.length==0) return (false);
for (Count=0; Count < InString.length; Count++)  {
	TempChar= InString.substring (Count, Count+1);
	if (RefString.indexOf (TempChar, 0)==-1)  
	return (false);
}
return (true);
}

function check()
{
if (document.form3.price.value=="")
{
alert("��*�ŵı�����д��");
return false;
}
if (isNumberString(document.form3.price.value,"1234567890.")!=1)
{
alert("���ֻ��Ϊ����!");
return false;
}
}
</script>
<%
nowtype=request("type")
if nowtype="" then nowtype=0
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../images/r_1.gif" alt="" /></td>
<td width="100%" background="../images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td>&nbsp;�������(��*�ŵ�Ϊ������)</td>
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
      <form name="form1">
      <tr>
        <td width="20%" height="30" align="right">���ͣ�</td>
        <td width="80%" class="category">
		<select name="type" onChange="form1.submit()">
          <option value="0"<%if nowtype="0" then%> selected="selected"<%end if%>>����</option>
          <option value="1"<%if nowtype="1" then%> selected="selected"<%end if%>>֧��</option>
        </select></td>
      </tr>	
	  </form>
	  <form name="form2">
	  <input type="hidden" name="type" value="<%=nowtype%>">
      <tr>
        <td align="right" height="30">�������ࣺ</td>
        <td class="category">
    <%
	sql="select * from money_bigclass where type="&nowtype&" order by id"
	set rs_bigclass=conn.execute(sql)
	if rs_bigclass.eof then
	%>
	<script language="javascript">
	  alert("�������������࣡")
	  window.location.href="../money/bigclass_add.asp"
	</script>
	<%
	  response.end
	end if
	nowbigclass=request("bigclass")
	if nowbigclass="" then
	  nowbigclass=rs_bigclass("id")
	end if
	%>
	  <select name="bigclass" onChange="form2.submit()">
        <%
	do while rs_bigclass.eof=false
	%>
        <option value="<%=rs_bigclass("id")%>"<%if trim(cstr(rs_bigclass("id")))=nowbigclass then%> selected="selected"<%end if%>><%=rs_bigclass("bigclass")%></option>
        <%
	  rs_bigclass.movenext
	loop
	%>
      </select>		  
		  </td>
      </tr>	
	  </form>
	  <form name="form3"> 
	  <input type="hidden" name="bigclass" value="<%=nowbigclass%>">
	  <input type="hidden" name="type" value="<%=nowtype%>">
      <tr>
        <td align="right" height="30">����С�ࣺ</td>
        <td class="category">
    <%
	sql="select * from money_smallclass where id_bigclass="&nowbigclass&" order by id"
	set rs_smallclass=conn.execute(sql)
	%>
	  <select name="smallclass">
	    <option value="0"></option>
        <%
	do while rs_smallclass.eof=false
	%>
        <option value="<%=rs_smallclass("id")%>"<%if trim(cstr(rs_smallclass("id")))=nowsmallclass then%> selected="selected"<%end if%>><%=rs_smallclass("smallclass")%></option>
        <%
	  rs_smallclass.movenext
	loop
	%>
      </select>  
		  </td>
      </tr>  
      <tr>
        <td align="right" height="30">��</td>
        <td class="category"><input type="text" name="price" style="width:100px" onKeyUp="value=value.replace(/[^\d.]/g,'')" onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d.]/g,''))">
		  <font color="#666666">Ԫ</font>
		  <font color="#ff0000">*</font>
		</td>
      </tr>	    
      <tr>
        <td align="right" height="30">�����ˣ�</td>
        <td class="category">
		  <%
		  if session("redboy_id")=1 then
			sql="select * from login order by id_zu,id"
			set rs_login=conn.execute(sql)
		  else
			sql="select * from login where id="&session("redboy_id")
			set rs_login=conn.execute(sql)	  
		  end if
		  if rs_login.eof then
		  %>
	      <script language="javascript">
	        alert("�������Ա����")
	        window.location.href="../system/user_add.asp"
	      </script>
		  <%
		  else
		  %>
		  <select name="id_login">
		  <%
		  do while rs_login.eof=false
		  %>
		    <option value="<%=rs_login("id")%>"<%if rs_login("username")=session("redboy_username") then%> selected="selected"<%end if%>><%=rs_login("username")%> (<%=rs_login("bianhao")%>)</option>
		  <%
		    rs_login.movenext
		  loop
		  %>
		  </select>
		  <%
		  end if
		  %>		</td>
      </tr>	
      <tr>
        <td align="right" height="30">��ע��</td>
        <td class="category">
          <textarea name="beizhu" cols="60" rows="3"></textarea>
        </td>
      </tr>	  	  
      <tr>
	    <td height="30">&nbsp;</td>
        <td class="category">
		  <input type="submit" value=" ȷ����� " onClick="return check()" class="button">&nbsp;&nbsp;&nbsp;&nbsp;
		  <input type="hidden" name="hid2" value="ok">
		  <input type="reset" value=" ������д " class="button">		</td>
      </tr>
	  </form>
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
</body>
</html>