<!--#include file="Conn.asp"-->
<!--#include file="Check.asp"-->
<!--#include file="Inc/Function.asp"-->
<!--#include file ="Inc/Date.asp"-->
<!-- #include file="../const.asp" --> 
<%
if fla86="0" and session("redboy_id")<>"1" then
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
xjcode=Request.QueryString("xjcode")
sqltext = "Select * From [XJRJZMX]"
If Request.QueryString("action") = "search" Then
selname = Request.Form("selname")
seltype = Request.Form("seltype")
adddate = Request.Form("adddate")
adddate2 = Request.Form("adddate2")
listid = Trim(Request.Form("listid"))

   If selname="" and seltype="" and adddate="" and listid="" and pjname="" Then
   sqltext = sqltext & " order by score,XJCODE,zt"
   Else
      sqltext = sqltext & " Where "
      If selname<>"" Then 
	     sqltext = sqltext & "score='"&selname&"' "
	  End if
          If seltype<>"" Then 
		     If selname<>"" Then sqltext = sqltext & "and "
		     sqltext = sqltext & "lb='"&seltype&"' "
		  End if

			 If adddate<>"" and adddate2<>"" Then
			    If seltype<>"" Then sqltext = sqltext & "and "
			    sqltext = sqltext & "CZRQ>=#"&adddate&"# and CZRQ<=#"&adddate2&"#"
			  End if

			 If adddate<>"" and adddate2="" Then
			    If seltype<>"" Then sqltext = sqltext & "and "
			    sqltext = sqltext & "CZRQ=#"&adddate&"# "
			  End if


				If listid<>"" Then
				   If adddate<>"" Then sqltext = sqltext & "and "
				   sqltext = sqltext & "XJNAME='"&listid&"' "
				End if
sqltext = sqltext & "  order by score,XJCODE,zt"
   End if

Else
	sqltext = sqltext & " order by score,XJCODE,zt"
End if
'response.write sqltext
if xjcode<>"" then
sqltext = "Select * From [XJRJZMX] where xjcode='"&xjcode&"' order by score,XJCODE,zt"
end if
rs.Open sqltext,Conn,1,3

MaxPerPage=2000
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
      <th width="9%">��־</th>
      <th width="18%">���</th>
      <th width="23%">��ʼ����</th>
      <th width="23%">��������</th>
      <th width="16%">��Ŀ</th>
      <th width="9%">&nbsp;</th>
    </tr>
    <tr height="22" valign="middle" align="center"> 
      <td><select name="selname" id="selname">
	   <option value="">ȫ��</option>
<%
Dim pa,pt
Set pa=Conn.Execute("Select score From [XJLB] group by score")
Do While Not pa.Eof
Response.write "<option value='" & pa(0) & "'>" & pa(0) & "</option>"
pa.MoveNext
Loop
pa.Close
Set pa=Nothing
%>
        </select></td>
      <td><select name="seltype" id="seltype">
	     <option value="">ȫ��</option>
<%
Set pt=Conn.Execute("Select lb From [XJLB] group by lb")
Do While Not pt.Eof
Response.write "<option value='" & pt(0) & "'>" & pt(0) & "</option>"
pt.MoveNext
Loop
pt.Close
Set pt=Nothing
%>
        </select></td>
      <td><input name="adddate" type="text" id="adddate" size="12" maxlength="12" readonly> 
        <input name="button" type="button" onclick="popUpCalendar(this, form1.adddate, 'yyyy-mm-dd')" value="ѡ��"></td>
      <td><input name="adddate2" type="text" id="adddate2" size="12" maxlength="12" readonly> 
        <input name="button" type="button" onclick="popUpCalendar(this, form1.adddate2, 'yyyy-mm-dd')" value="ѡ��"></td>

      <td><input name="listid" type="text" id="listid" size="20" maxlength="30"></td>

      <td><input type="submit" name="Submit" value="�ύ">
        <input type="reset" name="Submit2" value="����"></td>
    </tr>
  </table>
</form>
<table width="98%"  cellpadding="2" cellspacing="0" align="center" class="toptable grid" border="1"> 
<tr height="22" valign="middle" align="center"> 
    <th colspan="10">����ƾ֤ͳ��(��ϸ)</th>
  </tr>
  <tr> 
<td width="5%" class="category"> <div align="center">��־</div></td>
<td width="10%" class="category"> <div align="center">ƾ֤���</div></td>
<td width="10%" class="category"> <div align="center">��Ŀ����</div></td>
    <td width="15%" class="category"> <div align="center">��Ŀ����</div></td>
    <td width="8%" class="category"> <div align="center">���</div></td>
    <td width="10%" class="category"> <div align="center">����</div></td>
    <td width="5%" class="category"> <div align="center">��Ա</div></td>    
    <td width="10%" class="category"><div align="center">ժҪ</div></td>
    <td width="8%" class="category"><div align="center">���</div></td>    
    <td width="8%" class="category"><div align="center">״̬</div></td>        
  </tr>
<%
i=0
If rs.Eof Then
Response.Write "<tr><td colspan='8'>û�иü�¼!</td></tr>"
Else
ajhj=0
adhj=0
bjhj=0
bdhj=0

Do While Not rs.Eof

if rs("zt")="����" then
  If rs("score")="��"  then
ajhj=ajhj+cdbl(rs("xjmoney"))
else
adhj=adhj+cdbl(rs("xjmoney"))	
  end if	
end if  

if rs("zt")="δ��" then
  If rs("score")="��"  then
bjhj=bjhj+cdbl(rs("xjmoney"))
else
bdhj=bdhj+cdbl(rs("xjmoney"))	
  end if	
end if
%>  
   <tr>

<td><div align="center"><%=rs("SCORE")%></div></td>
    <td><div align="center"><%=rs("XJBH")%></a></div></td>
    <td><div align="center"><%=rs("XJCODE")%></div></td>
    <td><div align="center"><%=rs("XJNAME")%></div></td>
    <td><div align="center"><%=rs("xjmoney")%></div></td>
    <td><div align="center"><%=rs("CZRQ")%></div></td>
    <td><div align="center"><%=rs("CZR")%></div></td>
    <td><div align="center"><%=rs("REMARK")%></div></td>
    <td><div align="center"><%=rs("lb")%></div></td>
    <td><div align="center"><%=rs("zt")%></div></td>
  </tr>
<%
i=i+1 
if i >= MaxPerpage then exit do 
rs.MoveNext
Loop
End if
response.write "</p align='center'>"
response.write "��ǰ����跽�ϼ�:"
response.write ajhj
response.write "  ��ǰ��������ϼ�:"
response.write adhj
response.write "  ��ǰδ��跽�ϼ�:"
response.write bjhj
response.write "  ��ǰδ������ϼ�:"
response.write bdhj
Response.Write "</table>"

showpage()
Response.Write "</html>"
If Request.QueryString("action") = "del" Then
Dim dd,delid
delid=Request("id")
Set dd = Conn.Execute("Delete * From [PayList] Where ID ="&delid)
Response.Redirect "Pay_List3.asp"
End if
rs.Close
Conn.Close
Set rs = Nothing
Set Conn = Nothing
%>

