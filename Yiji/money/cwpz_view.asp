<!--#include file="Conn.asp"-->
<!--#include file="Check.asp"-->
<!--#include file ="Inc/Date.asp"-->
<%
Dim ye,ke,xe,j,Sql
Dim rss,ord_id,len_id,zero,i
Dim isadmin,kk_id,kk_na,kk
Dim cus
Dim pro_id,pro_na,pro_gg,pro_money
%>
<html>
<head>
<title>��������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style/style.css" rel="stylesheet" type="text/css">
<style>
body {
	background-color:#FFFFFF;
}
</style>
<script language="JavaScript" src="Image/js.js"></SCRIPT>
<script src="Js/AddOrder.js"></script> 
<script language="JavaScript" type="text/JavaScript">

function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
</script>
</head>
</head>
<script language="JavaScript">

function CheckNum(){
   if((event.keyCode<48||event.keyCode>57)&&event.keyCode!=46) event.returnValue=false; 
   }

</script>
<%
id = Request("id")
Set rs = Conn.Execute("Select * From [XJRJZ] Where ID="&id&"")
if not rs.eof then
%>
<body text="#000000">
<br>
<table width="98%"  align="center" cellpadding="0" cellspacing="0" class="toptable grid" border="1">
  <tr> 
    <td height="22"> <div align="left">
        </strong></font><strong>����[</strong>��˾����ƾ֤<strong>]</strong></div></td>
  </tr>
</table>
<br>
<table cellpadding="6" cellspacing="0" border="0" width="98%" class="tableBorder" align=center>
  <form action="cwpz_add2.asp" method=post id=form1 name=form1 onSubmit="return chk(this)">
    <tr bgcolor=ffffff> 
      <th height=25 align="center">��˾����ƾ֤</th>
    </tr>
    <tr bgcolor=ffffff> 
      <td class="category"> <table width="100%"  cellpadding="4" cellspacing="2" class="toptable grid" border="1">
          <tr> 
            <td>ƾ֤���:<input name="XJBH" type="text" id="XJBH" size="12" maxlength="12" value=<%=rs("XJBH")%>>
            	  ƾ֤����:<input name="CZRQ" type="text" id="CZRQ" size="12" maxlength="12" value=<%=rs("CZRQ")%> readonly>
            	  ������:<input name="CZR" type="text" id="CZR" size="8" maxlength="8" value=<%=rs("CZR")%>>
            	  �����:<input name="CZR" type="text" id="CZR" size="8" maxlength="8" value=<%=rs("shr")%>>
            	  ������:<input name="fj" type="text" id="fj" size="4" maxlength="6" value=<%=rs("fj")%>>
            	</td>
          <tr>
          </tr>
<td>
            	��ϸ������:<input name="spsl" type="text" id="spsl"  size="5" maxlength="10" value=<%=rs("mxs")%>> 
</td>

          
        </table>
        <br> 
<table width="100%"  cellpadding="4" cellspacing="1" class="toptable grid" border="1">
          <tr> 
            <td width="2%" class="category"> <div align="center">��Ŀ</div></td>
            <td width="4%" class="category"> <div align="center">���</div></td>
            <td width="9%" class="category"> <div align="center">��Ŀ����</div></td>
            <td width="8%" class="category"> <div align="center">��Ŀ����</div></td>
            <td width="5%" class="category"> <div align="center">���</div></td>
            <td width="10%" class="category"> <div align="center">ժҪ</div></td>
          <tr> 
<%
a=1
jhj=0
dhj=0
Set rs1 = Conn.Execute("Select * From [XJRJZMX] Where XJBH='"&rs("XJBH")&"'")
do while not rs1.eof
  If rs1("score")="��" then
jhj=jhj+cdbl(rs1("XJMONEY"))
else
dhj=dhj+cdbl(rs1("XJMONEY"))	
  end if	
%>
            <td colspan="7" class="category" id=upid> <table width="100%" border="0" cellpadding="2" cellspacing="0">
                <tr> 
                  <td width="2%"> <div align="center"><%=a%></div></td>
                  <td width="4%"> <div align="center"> 
                      <input name="score1" type="text" id="score1" size="4" maxlength="20" value=<%=rs1("score")%>>
</td>
                  <td width="10%"> <div align="center"> 
                      <input name="spid1" type="text" id="spid1" size="12" maxlength="20" value=<%=rs1("XJCODE")%>>
                    </div></td>
                  <td width="8%"> <div align="center"> 
                      <input name="pro_news1" type="text" id="pro_news1" size="25" maxlength="20" value=<%=rs1("XJNAME")%>>
                    </div></td>
                  <td width="5%"> <div align="center"></div>
                    <div align="center"> 
                      <input name="pro_sl1" type="text" id="pro_sl1" size="8" maxlength="20" value=<%=rs1("XJMONEY")%>>
                    </div></td>
                  <td width="10%"> <div align="center"> 
                      <input name="remark1" type="text" id="remark1" size="18" maxlength="60" value=<%=rs1("remark")%>>
                    </div></td>  
                </tr>
              </table>
             
              </td>
          </tr>
<%
a=a+1
rs1.movenext
loop
%> 
     <td colspan="4" width="10%" class="category"> <div align="center"></div></td>
     <td colspan="3" width="10%" class="category"> <div align="center">�跽�ϼ�:<%=jhj%> �����ϼƣ�<%=dhj%></div></td>     
        </table>
       </td>
    </tr>
<%
end if
set rs=nothing
set rs1=nothing
set conn=nothing
%>
  </form>
</table>
</html>