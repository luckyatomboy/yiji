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
<title><%=dianming%> - 客户退货</title>
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
if fla21="0" and session("redboy_id")<>"1" then
%>
<br><center><img src="../images/note.gif" align="absmiddle">&nbsp;<font color="#FF0000">你不具备此权限，请与管理员联系！</font></center>
<%  
  response.end
end if
%>

<%
if request("hid1")="ok" then
nowid_login=request("id_login")
set rs_login=conn.execute("select * from login where id="&nowid_login)
nowid_gys=request("id_gys")
nowgys=""
set rs_gys=conn.execute("select * from gys where id="&nowid_gys)
if rs_gys.eof=false then nowgys=rs_gys("company")
nowhuiyuan=request("huiyuan")
if nowhuiyuan="" then
  nowhuiyuan=0
end if
set rs_bianhao=conn.execute("select * from buy order by id desc")
if rs_bianhao.eof then
  nowbianhao=1000001
else
  nowbianhao=1000001+rs_bianhao("id")
end if

totalshulian=0
totalprice=0
totalprice2=0
for x=1 to maxproduit
nowku=request("ku"&x)
set rs_ku=conn.execute("select * from ku where id="&nowku)
nowhuohao=request("huohao"&x)
nowshulian=request("shulian"&x)
nowprice=request("price"&x)
if nowprice="" then
  nowprice=0
end if
if nowhuohao<>"单击选择产品" and nowshulian<>"" then
set rs_produit=conn.execute("select * from produit where huohao='"&nowhuohao&"'")
set rs=server.createobject("ADODB.RecordSet")
sql="select * from produit where huohao='"&nowhuohao&"' and id_ku="&nowku
rs.open sql,conn,1,3
if rs.eof then
  rs.addnew
  rs("id_bigclass")=rs_produit("id_bigclass")
  rs("id_smallclass")=rs_produit("id_smallclass")
  rs("title")=rs_produit("title")
  rs("huohao")=rs_produit("huohao")
  rs("id_ku")=nowku
  rs("shulian")=nowshulian
  rs("danwei")=rs_produit("danwei")
  rs("price")=rs_produit("price")
  rs("price2")=rs_produit("price2")
  rs("guige")=rs_produit("guige")
  rs("baojin")=rs_produit("baojin")
  rs("class")=rs_produit("class")
  rs("duihuan")=rs_produit("duihuan")
  rs("tichen_type")=rs_produit("tichen_type")
  rs("tichen")=rs_produit("tichen")
  rs("photo")=rs_produit("photo")
  rs("beizhu")=rs_produit("beizhu")
else
  rs("shulian")=rs("shulian")+nowshulian 
end if
rs.update
rs.close
sql="select bigclass from bigclass where id="&rs_produit("id_bigclass")
set rs_bigclass=conn.execute(sql)
sql="select smallclass from smallclass where id="&rs_produit("id_smallclass")
set rs_smallclass=conn.execute(sql)
if rs_smallclass.eof then
  smallclass=""
else
  smallclass=rs_smallclass(0)
end if
totalshulian=totalshulian+nowshulian
totalprice=totalprice+nowprice*nowshulian
totalprice2=totalprice2+rs_produit("price2")*nowshulian
sql="insert into buy(id_produit,bigclass,smallclass,title,huohao,id_ku,ku,shulian,guige,id_login,login,type,selldate,price,price2,id_huiyuan,photo,bianhao,id_gys,gys) values("&rs_produit("id")&",'"&rs_bigclass(0)&"','"&smallclass&"','"&rs_produit("title")&"','"&nowhuohao&"',"&rs_ku("id")&",'"&rs_ku("ku")&"',"&nowshulian&",'"&rs_produit("guige")&"',"&nowid_login&",'"&rs_login("username")&"',1,#"&date()&"#,"&nowprice&","&rs_produit("price2")&","&nowhuiyuan&",'"&rs_produit("photo")&"','"&nowbianhao&"',"&nowid_gys&",'"&nowgys&"')"
conn.execute(sql)
end if
next
if nowhuiyuan<>0 then
  conn.execute("update huiyuan set jifen=jifen-"&totalprice2&" where id="&nowhuiyuan)
end if
sql="insert into buy(shulian,id_login,login,type,selldate,price,price2,id_huiyuan,bianhao,zu,id_gys,gys) values("&totalshulian&","&nowid_login&",'"&rs_login("username")&"',1,#"&date()&"#,"&totalprice&","&totalprice2&","&nowhuiyuan&",'"&nowbianhao&"',true,"&nowid_gys&",'"&nowgys&"')"
conn.execute(sql)
%>
<script language="javascript">
alert("客户退货成功！")
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

function check1()
{
if (document.form1.huohao1.value=="单击选择产品")
{
alert("还没有选择产品！");
return false;
}
if (document.form1.shulian1.value=="")
{
alert("请输入数量！");
return false;
}
}
</script>

<script language="javascript">
function chg(objName, objName2, zekou)
{
	var obj;
	obj = document.getElementById(objName);
	obj2 = document.getElementById(objName2);
	temp = obj2.value*zekou
	//obj.value = temp.toFixed(2)
	obj.value = Math.round(temp*100)/100
}
</script>
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../images/r_1.gif" alt="" /></td>
<td width="100%" background="../images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td>&nbsp;客户退货(带*号的为必填项)</td>
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
	  <form name="form1" method="post">
      <tr>
        <td width="20%" align="right" height="30">选择产品：</td>
        <td width="80%" class="category">
          <table cellpadding="0" cellspacing="0" width="100%" border=0>
		    <%for x=1 to maxproduit%>
		    <tr id="cailiaohan<%=x%>"<%if x<>1 then%> style="display:none;"<%end if%>>
			  <td>
			    <input name="huohao<%=x%>" readonly onClick="JavaScript:window.open('produit2.asp?form=form1&field=huohao<%=x%>&field2=price<%=x%>&field3=hidprice<%=x%>&span1=showshulian<%=x%>','','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=yes,resizable=no,width=853,height=470,top=176,left=161');" style="width:80px" value="单击选择产品">
				数量: <input type="text" name="shulian<%=x%>" style="width:40px" onKeyUp="this.value=this.value.replace(/[^\d.]/g,'')" onafterpaste="this.value=this.value.replace(/[^\d.]/g,'')" value="1">
				单价: <input type="text" name="price<%=x%>" style="width:50px" onKeyUp="value=value.replace(/[^\d.]/g,'')" onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d.]/g,''))">
				<input type="hidden" name="hidprice<%=x%>">
				折扣：<select onChange="chg('price<%=x%>','hidprice<%=x%>',this.value)">
				  <option value="1">全价</option>
				  <option value="0.9">九折</option>
				  <option value="0.8">八折</option>
				  <option value="0.7">七折</option>
				  <option value="0.6">六折</option>
				  <option value="0.5">五折</option>
				  <option value="0.4">四折</option>
				  <option value="0.3">三折</option>
				  <option value="0.2">二折</option>
				  <option value="0.1">一折</option>
				</select>
				<%
				if session("redboy_id")=1 then
				  sql="select * from ku order by id"
				  set rs_ku=conn.execute(sql)
				else
				  sql="select * from ku where instr(login,'"&session("redboy_id")&",')>0 order by id"
				  set rs_ku=conn.execute(sql)	  
				end if
				if rs_ku.eof then
				%>
				<script language="javascript">
				  alert("没有属于你管理的仓库,请先添加仓库！")
				  window.location.href="../system/ku_add.asp"
				</script>
				<%
				end if
				%>
				仓库: <select name="ku<%=x%>">
				<%
				do while rs_ku.eof=false
				%>
			      <option value="<%=rs_ku("id")%>"<%if rs_ku("moren") then%> selected="selected"<%end if%>><%=rs_ku("ku")%></option>
				<%
				  rs_ku.movenext
				loop
				%>
				</select>							
				<%if x<>maxproduit then%><span onClick="cailiaohan<%=(x+1)%>.style.display=''" style="cursor:hand;">下一个产品</span><%end if%>
				<%if x=1 then%><font color="#ff0000">*</font><%end if%>
				<div id="showshulian<%=x%>"></div>
		      </td>
			</tr>		
			<%next%>		
		  </table>
		</td>
      </tr>
      <tr>	  
        <td height="30" align="right">选择会员：</td>
        <td class="category">
          <input name="huiyuan2" readonly onClick="JavaScript:window.open('../huiyuan/huiyuan1.asp?form=form1&field=huiyuan&field2=huiyuan2&field3=huiyuan3','','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=yes,resizable=no,width=853,height=470,top=176,left=161');" style="width:90px" value="单击选择会员"> <input type="text" name="huiyuan3" style="width:150px" readonly><input type="hidden" name="huiyuan"></td>	  		  
      </tr>	
	  <tr>
        <td align="right" height="30">供应商：</td>
        <td class="category">
		  <%
			    sql="select * from gys order by id"
			    set rs_gys=conn.execute(sql)
			    %>
			    <select name="id_gys">
				  <option value="0"></option>
			    <%
			    do while rs_gys.eof=false
			    %>
				  <option value="<%=rs_gys("id")%>"><%=rs_gys("company")%></option>
			    <%
				  rs_gys.movenext
			    loop
			    %>
			    </select>		
		</td>
      </tr>  
	  <tr>
        <td align="right" height="30">经办人：</td>
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
	        alert("请先添加员工！")
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
		  %>		
		</td>
      </tr>    
      <tr>
	    <td height="30">&nbsp;</td>
        <td class="category">
		  <input type="submit" value=" 确认退货 " onClick="return check1()" class="button">
		  <input type="hidden" name="hid1" value="ok">
		  <input type="reset" value=" 重新填写 " class="button">
		  </td>
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