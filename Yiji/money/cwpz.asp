<!--#include file="Conn.asp"-->
<!--#include file="Check.asp"-->
<!--#include file="Inc/Function.asp"-->
<!--#include file ="Inc/Date.asp"-->
<!-- #include file="../const.asp" --> 
<%
if fla83="0" and session("redboy_id")<>"1" then
%>
<br><center><img src="../images/note.gif" align="absmiddle">&nbsp;<font color="#FF0000">�㲻�߱���Ȩ�ޣ��������Ա��ϵ��</font></center>
<%  
  response.end
end if
%>
<%
Dim rs,news,sqltext
Dim id,lid,per,ptype,money,pj,menu,tm,io
Dim selname,seltype,listid,pjname,adddate
Set rs=Server.Createobject("Adodb.RecordSet")
sqltext = "Select XJBH,CZRQ,CZR,jehj,zt,ID,fj,shr From [XJRJZ] "
If Request.QueryString("action") = "search" Then
XJBH = Request.Form("XJBH")
CZR = Request.Form("CZR")
CZRQ = Request.Form("CZRQ")
zt = Request.Form("zt")
   If XJBH="" and CZR="" and CZRQ="" and zt="" Then
   sqltext = sqltext
   Else
      sqltext = sqltext & " Where "
      If XJBH<>"" Then 
	     sqltext = sqltext & "XJBH='"&XJBH&"'"
	  End if
      If CZR<>"" Then 
	     sqltext = sqltext & "CZR='"&CZR&"'"
	  End if
      If CZRQ<>"" Then 
	     sqltext = sqltext & "CZRQ=#"&CZRQ&"# "
	  End if	  
      If zt<>"" Then 
	     sqltext = sqltext & "zt='"&zt&"' "
	  End if	 

   End if
Else

sqltext = sqltext
End if
'response.write sqltext
rs.Open sqltext,Conn,1,3

MaxPerPage=200
text="0123456789"
rs.PageSize=MaxPerPage
for i=1 to Len(Request.QueryString("page"))
checkpage = Instr(1,text,mid(Request.QueryString("page"),i,1))
if checkpage=0 then
exit for
end if
next

If checkpage<>0 then
If NOT IsEmpty(Request.QueryString("page")) Then
CurrentPage=Cint(Request.QueryString("page"))
If CurrentPage < 1 Then CurrentPage = 1
If CurrentPage > rs.PageCount Then CurrentPage = rs.PageCount
Else
CurrentPage= 1
End If
If not rs.eof Then rs.AbsolutePage = CurrentPage end if
Else
CurrentPage=1
End if
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
</head>
<body text="#000000">
<script language="JavaScript" type="text/JavaScript">
function delpay()
{
   if(confirm("ȷ��Ҫɾ������"))
     return true;
   else
     return false;	 
}
</script>
<form name="form1" method="post" action="?action=search">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../images/r_1.gif" alt="" /></td>
<td width="100%" background="../images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td>&nbsp;�������</td>
	  <td align="right">&nbsp;</td>
    </tr>
  </table>
  <table width="98%"  cellpadding="0" cellspacing="0" align="center" class="toptable grid" border="1">
    <tr height="22" valign="middle" align="center"> 
     <th width="9%"  colspan="6" ><b><font color=FF0000>����ƾ֤����</b></th>
    <tr height="22" valign="middle" align="center"> 
      <th width="9%">״̬</th>
      <th width="23%">ƾ֤���</th>
      <th width="16%">��Ա</th>
      <th width="16%">����</th>
      <th width="9%">&nbsp;</th></td>
    </tr>
    <tr height="22" valign="middle" align="center"> 
      <td><select name="zt" id="zt">
	   <option value="">״̬</option>
<%
Dim pa,pt
Set pa=Conn.Execute("Select zt From [XJRJZ] group by zt")
Do While Not pa.Eof
Response.write "<option value='" & pa(0) & "'>" & pa(0) & "</option>"
pa.MoveNext
Loop
pa.Close
Set pa=Nothing
%>
        </select></td>

      <td><input name="XJBH" type="text" id="XJBH" size="26" maxlength="30"></td>
      <td><input name="CZR" type="text" id="CZR" size="16" maxlength="38"></td>
      <td><input name="CZRQ" type="text" id="CZRQ" size="16" maxlength="38"></td>      
      <td><input type="submit" name="Submit" value="�ύ">
        <input type="reset" name="Submit2" value="����"></td>
    </tr>
  </table>
</form>
<p align="center">
<a href="cwpz_add.asp"  class="category"><font color=FF0000><B>( ��������ƾ֤ )</B></a>
<table width="98%"  cellpadding="2" cellspacing="0" align="center" class="toptable grid" border="1"> 
<tr height="22" valign="middle" align="center"> 
    <th colspan="9">���в���ƾ֤</th>
  </tr>
  <tr> 
    <td width="10%" class="category"> <div align="center">ƾ֤���</div></td>
    <td width="10%" class="category"> <div align="center">ƾ֤����</div></td>
    <td width="6%" class="category"> <div align="center">��Ա</div></td>
    <td width="12%" class="category"> <div align="center">���</div></td>
    <td width="6%" class="category"> <div align="center">������</div></td>
    <td width="10%" class="category"><div align="center">״̬</div></td>
    <td width="10%" class="category"><div align="center">�����</div></td>
    <td width="8%" height="25" class="category"><div align="center">����</div></td>    
  </tr>
<%
i=0
If rs.Eof Then
Response.Write "<tr><td colspan='8'>û�иü�¼!</td></tr>"
Else
Do While Not rs.Eof
id = rs("id")
%>  
   <tr>
    <td height="25"><div align="center"><%=rs(0)%></div></td>
    <td><div align="center"><%=rs(1)%></div></td>
    <td><div align="center"><%=rs(2)%></div></td>
    <td><div align="center"><%=rs(3)%></div></td>
    <td><div align="center"><%=rs(6)%></div></td>
    <td><div align="center"><%=rs(4)%></div></td>
    <td><div align="center"><%=rs(7)%></div></td>   

    <td><div align="center"><a href="cwpz_view.asp?id=<%=id%>">��ʾ</a>
<%if rs(4)="δ��" then%>
    	 | <a href="?action=del&id=<%=id%>&bh=<%=rs(0)%>" onClick="return delpay();">ɾ��</a>
<%end if%>
    	</div></td>
  </tr>
<%
i=i+1 
if i >= MaxPerpage then exit do 
rs.MoveNext
Loop
End if
Response.Write "</table>"
showpage()
Response.Write "</html>"
If Request.QueryString("action") = "del" Then
Dim dd,delid
delid=Request("id")
bh=Request.QueryString("bh")


Set dd = Conn.Execute("Delete * From [XJRJZ] Where ID ="&delid)
Set dd = Conn.Execute("Delete * From [XJRJZMX] Where XJBH ='"&bh&"'")
Response.Redirect "cwpz.asp"
End if
rs.Close
Conn.Close
Set rs = Nothing
Set Conn = Nothing
%>

