<!--#include file="Conn.asp"-->
<!--#include file="Check.asp"-->
<!--#include file="Inc/Function.asp"-->
<!--#include file ="Inc/Date.asp"-->
<!-- #include file="../const.asp" --> 
<%
if fla82="0" and session("redboy_id")<>"1" then
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
sqltext = "Select XJCODE,XJNAME,XJYH,YHZH,YHLXR,XGDH,SCORE,LB,ID,qc From [XJLB] "
If Request.QueryString("action") = "search" Then
LB = Request.Form("LB")
XJNAME = Request.Form("XJNAME")
XJYH = Request.Form("XJYH")

   If LB="" and XJNAME="" and XJYH="" Then
   sqltext = sqltext
   Else
      sqltext = sqltext & " Where "
      If LB<>"" Then 
	     sqltext = sqltext & "LB='"&LB&"' "
	  End if
      If XJNAME<>"" Then 
	     sqltext = sqltext & "XJNAME='"&XJNAME&"' "
	  End if
      If XJYH<>"" Then 
	     sqltext = sqltext & "XJYH='"&XJYH&"' "
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
     <th width="9%"  colspan="6" class="toptable grid"><b><font color=FF0000>�����Ŀ����</b></th>
    <tr height="22" valign="middle" align="center"> 
      <th width="9%" >����</th>
      <th width="23%">����</th>
      <th width="16%">����</th>
      <th width="9%" >&nbsp;</th>
    </tr>
    <tr height="22" valign="middle" align="center"> 
      <td><select name="LB" id="LB">
	   <option value="">ȫ��</option>
<%
Dim pa,pt
Set pa=Conn.Execute("Select LB From [XJLB] group by LB")
Do While Not pa.Eof
Response.write "<option value='" & pa(0) & "'>" & pa(0) & "</option>"
pa.MoveNext
Loop
pa.Close
Set pa=Nothing
%>
        </select></td>

      <td><input name="XJNAME" type="text" id="XJNAME" size="26" maxlength="30"></td>
      <td><input name="XJYH" type="text" id="XJYH" size="16" maxlength="38"></td>
      <td><input type="submit" name="Submit" value="�ύ">
        <input type="reset" name="Submit2" value="����"></td>
    </tr>
  </table>
</form>
<p align="center">
<a href="cwlb_add.asp"  class="toptable grid"><font color=FF0000><B>( ���������Ŀ )</B></a>
<table width="98%"  cellpadding="2" cellspacing="0" align="center" class="toptable grid" border="1"> 
<tr height="22" valign="middle" align="center"> 
    <th colspan="10">���в����Ŀ(ע��:�ڳ�������������!)</th>
  </tr>
  <tr> 
    <td width="8%" class="category"> <div align="center">��Ŀ����</div></td>
    <td width="10%" class="category"> <div align="center">��Ŀ����</div></td>
    <td width="6%" class="category"> <div align="center">��������</div></td>
    <td width="12%" class="category"> <div align="center">�����˺�</div></td>
    <td width="8%" class="category"> <div align="center">��ϵ��</div></td>
    <td width="10%" class="category"><div align="center">��ϵ�绰</div></td>
    <td width="5%" class="category"> <div align="center">��־</div></td>
    <td width="8%" height="25" class="category"><div align="center">���</div></td>
    <td width="8%" height="25" class="category"><div align="center">�ڳ����</div></td>
    <td width="10%" height="25" class="category"><div align="center">����</div></td>    
  </tr>
<%
i=0
jhj=0
dhj=0
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
    <td><div align="center"><%=rs(4)%></div></td>
    <td><div align="center"><%=rs(5)%></div></td>
    <td><div align="center"><%=rs(6)%></div></td>
    <td><div align="center"><%=rs(7)%></div></td>
    <td><div align="center"><%=rs(9)%></div></td>
    <td><div align="center"><a href="cwlb_Edit.asp?id=<%=id%>">�޸�</a> | <a href="?action=del&id=<%=id%>" onClick="return delpay();">ɾ��</a></div></td>
  </tr>
<%
if rs(6)="��" then
jhj=jhj+cdbl(rs(9))
else
dhj=dhj+cdbl(rs(9))	
end if
i=i+1 
if i >= MaxPerpage then exit do 
rs.MoveNext
Loop
End if
%>
     <td colspan="3" width="10%" class="category"> <div align="center">�跽�ϼ�:<%=jhj%> �����ϼƣ�<%=dhj%></div></td>     
<%Response.Write "</table>"
showpage()
Response.Write "</html>"
If Request.QueryString("action") = "del" Then
Dim dd,delid
delid=Request("id")




Set dd = Conn.Execute("delete From [XJLB] Where ID ="&delid)
Response.Redirect "cwlb.asp"
End if
rs.Close
Conn.Close
Set rs = Nothing
Set Conn = Nothing
%>