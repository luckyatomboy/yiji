<!--#include file="Conn.asp"-->
<!--#include file="Check.asp"-->
<%
Dim proid
Dim cl,class_na
Dim id
id = Request("id")
proid = Request("pid")
'Set cl = Conn.Execute("Select LB From [XJLB] Where ID="&proid)
'if not cl.eof then
'class_na = cl(0)
'end if
'cl.Close
'Set cl=Nothing
%>
<HTML>
<HEAD>
<TITLE>��������</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<link href="../style/style.css" rel="stylesheet" type="text/css">
<style>
body {
	background-color:#FFFFFF;
}
</style>

<script language="JavaScript" src="Image/js.js"></SCRIPT>
<BODY text="#000000">
<SCRIPT language=javascript>
function CheckNum(){
   if((event.keyCode<48||event.keyCode>57)&&event.keyCode!=46) event.returnValue=false; 
   }
</SCRIPT>
<TABLE cellSpacing=1 cellPadding=0 width="98%" align=center class="toptable grid" border="1">
  <TBODY>
   <FORM name="Form2" action="?id=<%=id%>&pid=<%=proid%>&action=go" method="post">
    <TR> 
      <th align=middle colSpan=6 height=25>�����Ŀ�Ĳ�ѯ</th>
     
    </TR>
    <TR> 
      <td height=25 colSpan=6 align=middle class="category"> 
        <div align="center">��Ŀ���룺 
          <input name="bianhao" type="text" id="bianhao" size="10" maxlength="20">
          ��Ŀ���ƣ� 
          <input name="na" type="text" id="na" size="10" maxlength="30">
          ��� 
          <input name="from" type="text" id="from" size="12" maxlength="35">
          <br><br><input type="submit" name="Submit" value="�� ѯ">
        </div></td>
    </TR> </FORM>
    <TR class="category"> 
    <td align=middle bgcolor="#FFFFFF"> 
<%
Dim bianhao,na,from
Dim sqlstr,rr

bianhao = Trim(Request.Form("bianhao"))
na = Trim(Request.Form("na"))
from = Trim(Request.Form("from"))


if bianhao="" and na="" and from="" and 1=2 Then
Response.write "<script language='javascript'>alert('�������ѯ����!!');" & chr(13)
Response.write "window.document.location.href='javascript:history.go(-1)';</script>"
Else


Set rr=Server.Createobject("Adodb.RecordSet")
sqlstr = "Select XJCODE,XJName,SCORE,LB From [XJLB] Where 1=1 "
   If bianhao<>"" Then
      sqlstr = sqlstr & " and XJCODE='"&bianhao&"' "
   End if

           if na<>"" Then
              sqlstr = sqlstr & " and XJName like '%"&na&"%' "
           End if
              if from<>"" Then
                 sqlstr = sqlstr & " and LB like '%"&from&"%' "
              End if
  sqlstr = sqlstr & " Order By score"
''response.write sqlstr
rr.Open sqlstr,Conn,1,1
%>
        <table width="100%" class="toptable grid" border="1" cellpadding="0" cellspacing="0">
          <tr>
            <td><form name="form1" method="post" action="Order_Add.asp">
              <table width="100%"  cellpadding="4" cellspacing="1" class="toptable grid" border="1">
                <tr bgcolor="#FFFFFF"> 
                  <td height="10" colspan="6">&nbsp;</td>
                </tr>
                <tr bgcolor="#FFFFFF"> 
                  <td width="10%" height="25">
<div align="center">��Ŀ����</div></td>
                  <td width="20%"> 
                    <div align="center">��Ŀ����</div></td>
                  <td width="20%"> 
                    <div align="center">��Ŀ״̬</div></td>
                  <td width="25%"> 
                    <div align="center">���</div></td>
                  <td width="10%"><div align="center">ѡ��</div></td>
                </tr>
                <%
If rr.Eof Then
Response.Write "<tr><td colspan=7>û���κμ�¼!</td></tr>"
Else
Do While Not rr.Eof
%>
                <tr bgcolor="#FFFFFF"> 
                  <td height="25"><div align="center"><%=rr(0)%></div></td>
                  <td><div align="center"><%=rr(1)%></div></td>
                  <td><div align="center"><%=rr(2)%></div></td>
                  <td><div align="center"><%=rr(3)%></div></td>
                  <td><div align="center"> 
                      <input type="button" name="Submit3" value="ѡ��" onclick="insertsp('<%=rr(0)%>||<%=rr(1)%>||<%=rr(2)%>||<%=rr(3)%>');">
                    </div></td>
                </tr>
<%
rr.MoveNext
Loop
%>
                <tr bgcolor="#FFFFFF"> 
                  <td height="30" colspan="6"> <div align="center">
                      <input type="button" name="Submit2" value="�رմ���" onClick="javascript:window.close();">
                      <br>
                    </div></td>
                </tr>
<%
End if
rr.Close
Set rr=Nothing
Conn.Close
Set Conn=Nothing
%>
              </table>
              </form></td>
          </tr>
        </table>
<script>
function insertsp(spid){
 var Str;
 var Arr;
 Str = spid;
 Arr = Str.split("||");
 self.opener.form1.spid<%=id%>.value=Arr[0];
 self.opener.form1.pro_news<%=id%>.value=Arr[1];
self.opener.form1.score<%=id%>.value=Arr[2];
}
</script> 
<%

End if%>
    </td>
    </TR>
  </TBODY>
</TABLE>
</BODY>
</HTML>