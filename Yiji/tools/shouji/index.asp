  <%@ CODEPAGE = "936" %>
  
  <!--#include file="bbs.css" -->

<LINK href="style.css" type=text/css rel=stylesheet>
	<table width="100%" align=center border="0" cellspacing="0" cellpadding="0" class="tableBorder">
	  <tr> 
		<td>
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
		  <tr bgcolor=ffffff> 
			<th height=25 ><strong>�ֻ�������</strong></th>
		  </tr>
		</table>
		</td>
	  </tr>
	</table>
  <%
server.ScriptTimeout = 6000
   response.write("<title>��Խip-�ֻ���ַ��ѯ</title>")
  response.write("<body background=dw.gif>")
   myip = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
if myip = "" then myip = Request.ServerVariables("REMOTE_ADDR")
  if request.form("search")<>"" then
  fip=request.form("ip")
  c=fip
  feiyuedbq="db2.asp" 
   connstr="DBQ="+server.mappath(""&feiyuedbq&"")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
     set conn=server.createobject("ADODB.CONNECTION")
    conn.open connstr  
 '--------���ip��Ŀ����Ϊ�˸�����3λ�Ĳ�0��ͬʱҲ�ж�ip��ַ�Ƿ�Ϸ�--------
fip = Split(fip,".") 
if Ubound(fip)=3 then
For i=0 To Ubound(fip) 
if len(fip(i))=2 then
fip(i)="0"&fip(i)
elseif len(fip(i))=1 then
fip(i)="00"&fip(i)
end if
next
 oip=fip(0)&"."&fip(1)&"."&fip(2)&"."&fip(3)
 '--------�������ip�������ݿ��ip��ַ�Ƚϣ���ȡ��ַ,�������ֲ�ͬ������Ŀ����Ϊ�˵�ǰ����ip����ͬʱ���õ��ĵ�ַ�ȽϾ���--------
 sql1="select * from wry where  startip<='"&oip&"' and endip>='"&oip&"'and left(startip,3)='"&fip(0)&"' and  left(endip,3)='"&fip(0)&"' and  mid(startip,5,3)='"&fip(1)&"' and  '"&fip(1)&"'=mid(endip,5,3) and  mid(startip,9,3)='"&fip(2)&"' and  '"&fip(2)&"'=mid(endip,9,3)" 
 set rs1=conn.execute(sql1)
 if not rs1.eof then
 a=rs1("country")
 b=rs1("local")
  set rs1=nothing
else
 sql="select * from wry where  startip<='"&oip&"' and endip>='"&oip&"'and left(startip,3)='"&fip(0)&"' and  left(endip,3)='"&fip(0)&"' and  mid(startip,5,3)='"&fip(1)&"' and  '"&fip(1)&"'=mid(endip,5,3)" 
 set rs=conn.execute(sql)
 if not rs.eof then
 a=rs("country")
 b=rs("local")
  set rs=nothing
else
sql2="select * from wry where startip<='"&oip&"' and endip>='"&oip&"' and left(startip,3)='"&fip(0)&"' and  left(endip,3)='"&fip(0)&"'"
    set rs2=conn.execute(sql2)
    if not rs2.eof then
 a=rs2("country")
 b=rs2("local")
set rs2=nothing
else
sql3="select * from wry where startip<='"&oip&"' and endip>='"&oip&"'"
    set rs3=conn.execute(sql3)
    if not rs3.eof then
 a=rs3("country")
 b=rs3("local")
set rs3=nothing
else
a="�Ҳ�����ַ�������֪�������뱾վ��ϵ��лл��"
end if
end if
  end if 
end if 
  else
  a="����:�������ip��ַ�Ƿ������������룡"
  end if  
    
   %>

    <div align="center">
      <center>
    <table border="1" cellpadding="4" style="border-collapse: collapse" bordercolor="<%=bgcolor%>" width="500" id="AutoNumber1">
    <tr>
      <td width="100%" colspan="2" align="center" background="b1.gif" height="22" class=tdc1>�������:</td>
    </tr>
    <tr>
      <td width="28%" align="center" height="22" bgcolor=<%=tcolor2%> class=tdc>��Ҫ������IP:</td>
      <td width="75%" align="center" height="22" bgcolor=<%=tcolor2%> class=tdc><%=c%></td>
    </tr>
    <tr>
      <td width="28%" align="center" height=22 bgcolor=<%=tcolor2%> class=tdc>���ڵ�ַ:</td>
      <td width="75%" align="center" height=22 bgcolor=<%=tcolor2%> class=tdc><%=a%> <%=b%></td>
    </tr>
  </table>

      </center>
  </div>
<p>
   <%end if%>
     <%if request.form("searchsj")<>"" then
  d=request.form("num1")
  if len(d)>=7 then
  c=left(d,7)
  feiyuedbq="db1.asp" 
   connstr="DBQ="+server.mappath(""&feiyuedbq&"")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
     set conn=server.createobject("ADODB.CONNECTION")
    conn.open connstr  

  sql="select * from mod where  start<="&c&" and end>="&c&" and mid(start,3,1)="&mid(c,3,1)&" and  mid(end,3,1)="&mid(c,3,1)&" and mid(start,4,1)="&mid(c,4,1)&" and  mid(end,4,1)="&mid(c,4,1)&" and mid(start,5,1)="&mid(c,5,1)&" and  mid(end,5,1)="&mid(c,5,1)&" and mid(start,6,1)="&mid(c,6,1)&" and  mid(end,6,1)="&mid(c,6,1)&"  " 
 set rs=conn.execute(sql)
 if not rs.eof then
 a=rs("city")
 b=rs("city1")
  e=rs("url")
  set rs=nothing
else
  sql="select * from mod where  start<="&c&" and end>="&c&" and mid(start,3,1)="&mid(c,3,1)&" and  mid(end,3,1)="&mid(c,3,1)&" and mid(start,4,1)="&mid(c,4,1)&" and  mid(end,4,1)="&mid(c,4,1)&" and mid(start,5,1)="&mid(c,5,1)&" and  mid(end,5,1)="&mid(c,5,1)&"  " 
 set rs=conn.execute(sql)
 if not rs.eof then
 a=rs("city")
 b=rs("city1")
  e=rs("url")
  set rs=nothing
else
  sql="select * from mod where  start<="&c&" and end>="&c&" and mid(start,3,1)="&mid(c,3,1)&" and  mid(end,3,1)="&mid(c,3,1)&" and mid(start,4,1)="&mid(c,4,1)&" and  mid(end,4,1)="&mid(c,4,1)&"   " 
 set rs=conn.execute(sql)
 if not rs.eof then
 a=rs("city")
 b=rs("city1") 
 e=rs("url")
  set rs=nothing
  else
  sql="select * from mod where  start<="&c&" and end>="&c&" and mid(start,3,1)="&mid(c,3,1)&" and  mid(end,3,1)="&mid(c,3,1)&"  " 
 set rs=conn.execute(sql)
 if not rs.eof then
 a=rs("city")
 b=rs("city1")
  e=rs("url")
  set rs=nothing
  else
  sql="select * from mod where  start<="&c&" and end>="&c&"" 
 set rs=conn.execute(sql)
 if not rs.eof then
 a=rs("city")
 b=rs("city1")
 e=rs("url")
  set rs=nothing
else
a="�Ҳ�����ַ�������֪�������뱾վ��ϵ��лл��"
end if
end if
  end if  
  end if
  end if
  else
  a="����:��ȷ����������Ǵ���7λ�����֣�"
  end if  
   %>
    <div align="center">
      <center>
    <table border="1" cellpadding="4" style="border-collapse: collapse" bordercolor="<%=bgcolor%>" width="500" id="AutoNumber1">
    <tr>
      <td width="100%" colspan="2" align="center" background="b1.gif" height="22" class=tdc1>�������:</td>
    </tr>
    <tr>
      <td width="28%" align="center" height="22" bgcolor=<%=tcolor2%> class=tdc>��Ҫ�������ֻ�����:</td>
      <td width="75%" align="center" height="22" bgcolor=<%=tcolor2%> class=tdc><%=d%></td>
    </tr>
    <tr>
      <td width="28%" align="center" height=22 bgcolor=<%=tcolor2%> class=tdc>���ڵ�ַ:</td>
      <td width="75%" align="center" height=22 bgcolor=<%=tcolor2%> class=tdc><%=a%> <%=b%> </td>
    </tr>
    <tr>
      <td width="28%" align="center" height=22 bgcolor=<%=tcolor2%> class=tdc>���ź�����:</td>
      <td width="75%" align="center" height=22 bgcolor=<%=tcolor2%> class=tdc><%=e%></td>
    </tr>
  </table>

      </center>
  </div>
<p>
   <%end if
     feiyuedbq="db2.asp" 
   connstr="DBQ="+server.mappath(""&feiyuedbq&"")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
     set conn=server.createobject("ADODB.CONNECTION")
    conn.open connstr 
   '--------���ip��Ŀ����Ϊ�˸�����3λ�Ĳ�0��ͬʱҲ�ж�ip��ַ�Ƿ�Ϸ�--------
myip1 = Split(myip,".") 
if Ubound(myip1)=3 then
For i=0 To Ubound(myip1) 
if len(myip1(i))=2 then
myip1(i)="0"&myip1(i)
elseif len(myip1(i))=1 then
myip1(i)="00"&myip1(i)
end if
next
 myip2=myip1(0)&"."&myip1(1)&"."&myip1(2)&"."&myip1(3)
 '--------�������ip�������ݿ��ip��ַ�Ƚϣ���ȡ��ַ,�������ֲ�ͬ������Ŀ����Ϊ�˵�ǰ����ip����ͬʱ���õ��ĵ�ַ�ȽϾ���--------
 sql1="select * from wry where  startip<='"&myip2&"' and endip>='"&myip2&"'and left(startip,3)='"&myip1(0)&"' and  left(endip,3)='"&myip1(0)&"' and  mid(startip,5,3)='"&myip1(1)&"' and  '"&myip1(1)&"'=mid(endip,5,3) and  mid(startip,9,3)='"&myip1(2)&"' and  '"&myip1(2)&"'=mid(endip,9,3)" 
 set rs1=conn.execute(sql1)
 if not rs1.eof then
 a=rs1("country")
 b=rs1("local")
  set rs1=nothing
else
 sql="select * from wry where  startip<='"&myip2&"' and endip>='"&myip2&"'and left(startip,3)='"&myip1(0)&"' and  left(endip,3)='"&myip1(0)&"' and  mid(startip,5,3)='"&myip1(1)&"' and  '"&myip1(1)&"'=mid(endip,5,3)" 
 set rs=conn.execute(sql)
 if not rs.eof then
 a1=rs("country")
 b1=rs("local")
  set rs=nothing
else
sql2="select * from wry where startip<='"&myip2&"' and endip>='"&myip2&"' and left(startip,3)='"&myip1(0)&"' and  left(endip,3)='"&myip1(0)&"'"
    set rs2=conn.execute(sql2)
    if not rs2.eof then
 a1=rs2("country")
 b1=rs2("local")
set rs2=nothing
else
sql3="select * from wry where startip<='"&myip2&"' and endip>='"&myip2&"'"
    set rs3=conn.execute(sql3)
    if not rs3.eof then
 a1=rs3("country")
 b1=rs3("local")
set rs3=nothing
else
a="�Ҳ�����ַ�������֪�������뱾վ��ϵ��лл��"
end if
end if
  end if 
end if 
else
a1="��Ǹ���Ҳ�����ַ��"
end if %>

    <div align="center">
      <center>
    <table border="1" cellpadding="4" style="border-collapse: collapse" bordercolor="<%=bgcolor%>" width="500" id="AutoNumber1">
    <tr>
      <td width="100%" align="center" background="b1.gif" height="22" class=tdc1 colspan="2">
      ����ĳһ IP ��ַ�ĵ���λ��</td>
    </tr>
    <tr>
      <td width="28%" align="center" height=22 bgcolor=<%=tcolor2%> class=tdc>
    
  ������Ҫ������IP��ַ:
</td>
      <form method="post" action="index.asp">   
    <td width="75%" align="center" height=22 bgcolor=<%=tcolor2%> class=tdc>
    <input type="text" name="ip" size="20" class=tdc> 
    <input type="hidden" name="search" size="20" value=search>
    <input type="submit" value="�� ѯ" name="B1" class=bdtj>
    
 
 </td> </form>
    </tr>
       <tr>
      <td width="28%" align="center" height=22 bgcolor=<%=tcolor2%> class=tdc>
    
  �㵱ǰ���ڵ�IP:
</td>
      
    <td width="75%" align="center" height=22 bgcolor=<%=tcolor2%> class=tdc>
   
    <%=myip%>  ���ԣ�<%=a1%> <%=b1%>
         
 
 </td> 
    </tr>
  </table>

      </center>
  </div>
      <div align="center">
      <center>
    <table border="1" cellpadding="4" style="border-collapse: collapse" bordercolor="<%=bgcolor%>" width="500" id="AutoNumber1">
    <tr>
      <td width="100%" align="center" background="b1.gif" height="22" class=tdc1 colspan="2">
      �����ֻ������Ӧ�ĵ���λ��</td>
    </tr>
    <tr>
      <td width="28%" align="center" height=22 bgcolor=<%=tcolor2%> class=tdc>
    
  ������Ҫ�����ĺ���:
</td>
      <form method="post" action="index.asp">   
    <td width="75%" align="center" height=22 bgcolor=<%=tcolor2%> class=tdc>
    <input type="text" name="num1" size="20" class=tdc> 
    <input type="hidden" name="searchsj" size="20" value=search>
    <input type="submit" value="�� ѯ" name="B1" class=bdtj>
    
 
 </td> </form>
    </tr>
  </table>

      </center>
  </div>