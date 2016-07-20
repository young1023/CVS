<!--#include file="include/SQLConn.inc" -->
<!--#include file ="js/OVERLIB.JS" -->
<!--#include file ="js/OVERLIB_MINI.JS" -->
<!--#include file ="js/select_date.JS" -->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>Shell CVS Administration</title>
<link rel="stylesheet" type="text/css" href="include/hse.css" />
<SCRIPT language=JavaScript>
<!--

function gtpage(what)
{
document.fm1.pageid.value=what;
document.fm1.action="dbsize1.asp"
document.fm1.submit();
}

function findenum()
{
document.fm1.action="dbsize1.asp"
document.fm1.submit();
}
//-->
</script>

</head>

<body leftmargin="0" topmargin="0" >
<div id="overDiv" style="position:absolute; visibility:hidden; z-index:1000;"></div>
<!--#include file="include/header.inc" -->
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td colspan="3" class="NavyBlue" height="3"></td>
  </tr>
  <tr>
    <td colspan="3" class="NavyBlue">
    <table border="0" width="100%" cellspacing="0" cellpadding="0" class="TitleBar">
      <tr>
        <td width="100%">
        <table border="0" width="100%" cellspacing="0" cellpadding="0" class="TitleBar">
            <tr>
              <td><div align="center"><font class="Head" style="font-size: 13px">Shell CVS Administration</font></div></td>
            </tr>
            </table>
          </td>
          <td nowrap="true">
          </td>
        </tr>
      </table>
      </td>
    </tr>
    <tr>
      <td colspan="3" class="NavyBlue" height="1"></td>
    </tr>
    <tr>
      <td colspan="3" height="1"></td>
    </tr>
  </table>
 




                    <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#E6EBEF" class="normal">
                     <tr> 
                          
                        <td height="28" align="center"><font color="#FF6600"><b>Database Size</b></font></td>
                        </tr>
                        <tr> 
                          <td valign="top" align="center">
                            <form name=fm1 method=post>
                            <table width="99%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" class="normal">
			 
                              <tr> 
                                <td height="28"> 
<% 
' By getting the pageid, the system knows which page to go.

		pageid=trim(request.form("pageid"))
		if pageid="" then
		  pageid=1
		end if
    
' Start the queries
      
    fsql = "SELECT D.name, F.Name AS FileType, F.physical_name AS PhysicalFile, "
    fsql = fsql & "F.state_desc AS OnlineStatus, "
    fsql = fsql & "(F.size*8)/1000  AS FileSize, "
    fsql = fsql & "CAST((F.size*8)/1000000 AS VARCHAR(32)) + 'GB' as SizeInGBytes "
    fsql = fsql & "FROM sys.master_files F INNER JOIN sys.databases D ON "
    fsql = fsql & "D.database_id = F.database_id ORDER BY D.name"

    Set frs = Conn.Execute(fsql)


        'response.write fsql
        'response.end

' Setting the page

        set frs=createobject("adodb.recordset")
		frs.cursortype=1
		frs.locktype=1
        frs.open fsql,conn

       if frs.RecordCount=0 then
           response.write "<font color=red>No Record</font>"
       else
          findrecord=frs.recordcount
          response.write "Total <font color=red>"&findrecord&"</font> Records ;"

         frs.PageSize = 100

' Call the function to count the page.

         
         end if
	     'response.write "&nbsp;&nbsp;<input type='text' name='findnum' size='13' value='"&findnum&"'>"
		 'response.write "&nbsp;&nbsp;<input type='button' value='   Search   ' onClick='findenum();' class='common'>"
	   
%>


   </td>
      </tr>
         <tr> 
            <td valign="top" height="28">

<% ' ----------------------------------------------------------------
   ' Main table of the page, change content here for different module 
   ' ----------------------------------------------------------------
%>


   <table border="0" align=center cellpadding="1" width="99%" cellspacing="1" class="normal">
     <tr bgcolor="#DFDFDF">
<td width="7%" height="28">Database Name</td>
<td >File Type</td>

<td width="5%" height="28">Physical Path</td>

<td width="10%" height="28">Size in MB</td>

<td >Size In GB</td>

<td width="10%" height="28">Online Status</td>

</tr>
                                    <%



 Total =0
 if frs.recordcount>0 then
  frs.AbsolutePage = pageid
  do while (frs.PageSize-i)
   if frs.eof then exit do
   
%>
   <tr>

<td ><% = frs("Name")%></td>
<td  height="28">
<% = frs("FileType") %>
</td>
<td  height="28"><% = frs("PhysicalFile") %>
</td>

<td ><% = frs("FileSize") %>
</td>

<td ><% = frs("SizeInGBytes") %>
</td>

<td height="28"><% = frs("OnlineStatus") %>
</td>

</tr>
<%
   Total = Total + frs("FileSize")

   frs.movenext
  loop
 end if
  %>

<tr bgcolor="#DFDFDF">

<td colspan="4" align="right">Total:</td>

<td width="10%" height="28"><% = formatnumber(Total/1000,2) %> GB</td>

<td ><% = formatnumber(Total/300,2) %>% Full</td>

</tr>


                                  </table>
                                </td>
                              </tr>
                              <tr> 
                                <td align="right" height="28"> 

                                 
                                </td>
                              </tr>
                              <tr> 
                                <td height="28" align="center"> 

<%
  
		      frs.close
			  set frs=nothing
			  conn.close
			  set conn=nothing
%>
                                                         
 </td>
                              </tr>
                              <tr> 
                                <td valign="top"></td>
                              </tr>
                            </table>
                          </form>

</td>
                              </tr>
                            </table>

              </body>

              </html>