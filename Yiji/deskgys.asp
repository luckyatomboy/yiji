<HTML><HEAD><TITLE>ϵͳ����</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="style/Style.css" type=text/css rel=stylesheet>
<STYLE type=text/css>
BODY {
	MARGIN: 0px 0px; 
	BACKGROUND-COLOR: #ffffff;
}
</STYLE>
</HEAD>
<!--#include file="conn.asp"-->
<!--#include file="info/include/tools.asp"-->
<%
oabusyname=Session("EmpName")
oabusyusername=Session("ManageName")
oabusyuserdept=Session("DeptName")
session("id")=Session("EmpId")

%>
<SCRIPT language="JavaScript">
	function MM_Do(id,action){
		if(action=="delete"){
			if(confirm("ȷ��Ҫɾ��������Ϣ��\nע�⣺ɾ�����ָܻ���")){
				document.frm.action="mail/DeleteInbox.asp";
			}
		}else if(action=="brow"){
			document.frm.action="mail/BrowInbox.asp";
		}

		document.frm.MailId.value=id;
		document.frm.submit();
	}
</SCRIPT>

<script language="javascript">
  function OpenSmallWindows(strURL)
  {
     window.open (strURL,"_blank",
     "status=no,resizable=0,toolbar=no,menubar=no,scrollbars=no,width=130,height=300,left=600,top=125");
  }
function openwinfun(hrefvalue,winname,heightvalue,widthvalue)
{
	var statusstr;
	statusstr='location=no,height='+heightvalue+', width='+widthvalue+', toolbar=no, menubar=no, scrollbars=yes, resizable=no, location=no, status=no'
	window.open(hrefvalue,winname,statusstr);
}

</script>
<BODY>
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="images/r_1.gif" alt="" /></td>
<td width="100%" background="images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td>&nbsp;����</td>
	  <td align="right">&nbsp;</td>
    
    </tr>
  </table>
  </table>






    <TD vAlign=top height="50%">
      <TABLE class=topdashed cellSpacing=0 cellPadding=0 width="100%" 
        border=0><TBODY>
        <TR>
          <TD align=right height=23>
            <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
              <TBODY>
              <TR>
                <TD>��</TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
      <TABLE cellSpacing=0 cellPadding=0 width="90%" align=center border=0>
        <TBODY>
        <TR>
          <TD height=10>
            <DIV align=right></DIV></TD></TR>
        <TR>
          <TD height=25>��</TD></TR></TR></TBODY></TABLE></TD></TR>
  <TR>
<p></p>
  <form name="frm" method="post" >
    
<TABLE  cellSpacing=0 cellPadding=0 width="98%" align=center class="toptable grid" border="1">
  <TBODY>
    <TR> 
      <TH class="category" colSpan=2 height=22>�� �� �� ��</TH>

    <TR> 
      <TD class=forumRow width="50%" 
      height=23>&nbsp;&nbsp;<a href="tools/rl/cal.htm">��������</a></TD>
      <TD class=forumRow width="50%"><a href="tools/rl/time.htm">&nbsp;&nbsp;����ʱ��</a></TD>
    </TR>
    <TR> 
      <TD class=forumRow width="50%" 
      height=23>&nbsp;&nbsp;<a href="tools/yzqh/default.asp">��������</a></TD>
      <TD class=forumRow width="50%">&nbsp;&nbsp;<a href="tools/shouji/index.asp">�ֻ�����</a></TD>
    </TR>
    <TR>     	
      <TD class=forumRow width="50%" 
      height=23>&nbsp;&nbsp;<a href="#" onclick ="javascript:openwinfun('tools/rl/js1.html','qymlwin',400,550);">���߼�����</a></TD>
    <TD class=forumRow width="50%">&nbsp;&nbsp;</TD>
    </TR>
    <TR>     	

  </TBODY>
</TABLE>
      <TABLE cellSpacing=0 cellPadding=0 width="90%" align=center border=0>
        <TBODY>
        <TR>
          <TD align=right height=25>��Ʒ�����������Ƽ�<A href="http://www.redoa.com/" 
            target=_blank><SPAN class=STYLE1>(www.redoa.com)</SPAN> </A></TD></TR>
        <TR>
          <TD align=right height=30>����֧��QQ��52721918 QQ:267478238 Email��hxsd@citiz.net redoa@163.com</TD></TR></TBODY></TABLE></TD></TR>
  <TR>
    <TD>
      <DIV 
      style="COLOR: #003300; TEXT-ALIGN: right">----------------------------------------------------------------------------------------------------------------------------</DIV>
      <DIV style="HEIGHT: 30px; TEXT-ALIGN: right">Copyright 
      (c) 2007-2008 <A href="http://www.redoa.com/" target=_blank><FONT 
      color=#ff6600>www.redoa.com</FONT></A>. All Rights Reserved . 
  </DIV></TD></TR></TBODY></TABLE></BODY></HTML>