<!--#include file="include/SQLConn.inc" -->
<!--#include file ="js/OVERLIB.JS" -->
<!--#include file ="js/OVERLIB_MINI.JS" -->
<!--#include file ="js/select_date.JS" -->

<% 

' check which page is it
pageid=request("pageid")


From_Date   = Request.Form("From_Date")

if From_Date = "" then
   From_Date = formatdatetime(now(),2) 
end if

To_Date     = Request.Form("To_Date")

if To_Date = "" then
   To_Date = formatdatetime(now(),2) 
end if

Print_Excel = Request.Form("Print_Excel")

if Print_Excel = "" Then
   Print_Excel = "N"
End If

Excel_Type  = Request.Form("Excel_Type")
whatdo      = Request.Form("Whatdo")

If whatdo = "printexcel" Then


      set rs = server.createobject("adodb.recordset")
      'response.write  ("Exec PrintExcel '"&From_Date&"', '"&To_Date&"', '"&Excel_Type&"' ") 
	  rs.open ("Exec PrintExcel '"&From_Date&"', '"&To_Date&"', '"&Excel_Type&"' ") ,  conn,3,1




End if


%>
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
document.fm1.action="wk_rpt1.asp"
document.fm1.submit();
}

function findenum()
{
document.fm1.action="wk_rpt1.asp"
document.fm1.submit();
}

function exportExcel()
{
document.fm1.action="wk_rpt_excel1.asp"
document.fm1.submit();
}

function updateExcel()
{
document.fm1.action="wk_rpt1.asp"
document.fm1.whatdo.value="printexcel"
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
  <table width="100%" border="0" cellspacing="0" cellpadding="0" height="60%">
    <tr>
      <td width="1" class="HSEBlue"></td>
      <td valign="top">
      <table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="4" class="HSEBlue"></td>
        </tr>
        <tr valign="top">
          <td height="25"><img src="images/Curve.gif" width="22" height="16" /></td>
        </tr>
        <tr valign="top">
          <td height="100%" align="middle">


 <table width="100%" border="0" cellpadding=2 cellspacing="0" bgcolor="#FFFFFF" height="100%">
                <tr>
                  <td valign="top" align="center" bgcolor="#E6EBEF">
                    <table width="100%" border="0" cellpadding="2" cellspacing="0" bgcolor="#E6EBEF" class="normal">
                     <tr> 
                          
                        <td height="28" align="center"><font color="#FF6600"><b>
Weekly Report</b></font></td>
                        </tr>
                        <tr> 
                          <td valign="top" align="center">
                            <form name=fm1 method=post>
                            <table width="100%" border="0" cellpadding="4" cellspacing="0" bgcolor="#FFFFFF" class="normal">
			 
                              <tr> 
                                <td height="28"> 
<% 
' By getting the pageid, the system knows which page to go.

		pageid=trim(request.form("pageid"))
		if pageid="" then
		  pageid=1
		end if
        findnum=replace(trim(request.form("findnum")),"%","％")
        findnum=replace(findnum,"'","''")
   
' Start the queries
         
      set frs = server.createobject("adodb.recordset")
     'response.write  ("Exec WeeklyReport '"&From_Date&"', '"&To_Date&"', '"&Print_Excel&"', '"&Excel_Type&"' ") 
	  frs.open ("Exec WeeklyReport '"&From_Date&"', '"&To_Date&"', '"&Print_Excel&"', '"&Excel_Type&"' ") ,  conn,3,1

%>


Date From:
<input type="text" name="From_Date" size="10" value="<% = From_Date %>">
<a href="javascript:show_calendar('fm1.From_Date');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a full year pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;"><img src="images/show-calendar.gif" width=24 height=22 border=0></a>
To Date:
<input type="text" name="To_Date" size="10" value="<% = To_Date %>">
<a href="javascript:show_calendar('fm1.To_Date');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a full year pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;"><img src="images/show-calendar.gif" width=24 height=22 border=0></a>
Print Excel: 
	<select size="1" name="Print_Excel" class="common">
			<option value="N" <%If Print_Excel="N" Then%>Selected<%End If%>>No</option>
			<option value="Y" <%If Print_Excel="Y" Then%>Selected<%End If%>>Yes</option>

	</select>
Excel Type :
<input type="text" name="Excel_Type" size="4" value="<% = Excel_Type %>">
<input type="button" value="   Search   " onClick="findenum();" class="common">

<% if From_Date <> "" Then %>

<input type="button" value="   Export to Excel   " onClick="exportExcel();" class="common">
<input type="button" value="   Update all to print Excel   " onClick="updateExcel();" class="common">

<% End If %>
   </td>
      </tr>
         <tr> 
            <td valign="top" height="28">

<% ' ----------------------------------------------------------------
   ' Main table of the page, change content here for different module 
   ' ----------------------------------------------------------------
%>


   <table border="0" align=center cellpadding="1" width="100%" cellspacing="1" class="normal">
     <tr bgcolor="#DFDFDF">

<td height="28">From Date</td>
<td height="28">To Date</td>
<td height="28">(Sum) Total / Amount (Sum)</td>
<td  height="28">Print Excel</td>
<td  height="28">Excel Type</td>
<td  height="28">Station</td>
<td >Ship To Code</td>
<td  height="28">Sold to Code</td>
</tr>
                                    <%



 i=0
 
 
  do while not frs.EoF
  
%>
   <tr>

<td align=center width="95" height="28"><% = frs("Doc Date")%></td>
<td  height="28"><% = frs("To Date") %>
</td>

<td height="28"><% = frs("Amount") %>
</td>

<td  height="28"><% = frs("Print Excel") %>
</td>

<td  height="28">
<%
   If frs("Excel_Type") <> "" then

     Response.Write frs("Excel_Type")

    End If
%>
</td>

<td >
<% = frs("Station") %>
</td>

<td >
<% = frs("ShipToCode") %>
</td>

<td>
<% = frs("SoldToCode") %> 
</td>

</tr>
<%
   
   frs.movenext
  loop

 
  %>

                                  </table>

<%


      ' Calculate the total

      set frs1 = server.createobject("adodb.recordset")
      'response.write  ("Exec WeeklyReport_Total '"&From_Date&"', '"&To_Date&"', '"&Print_Excel&"', '"&Excel_Type&"' ") 
      frs1.open ("Exec WeeklyReport_Total '"&From_Date&"', '"&To_Date&"', '"&Print_Excel&"', '"&Excel_Type&"' ") ,  conn,3,1


%>



<br>


 <table border="0" cellpadding="1" width="40%" cellspacing="1" class="normal">

     <tr bgcolor="#DFDFDF">

<%

    ' Total
    do while not frs1.EoF
 
  %>

  <tr>

<td>Total Amount of Excel Type: <% = frs1("Excel_Type") %>
</td>

<td height="28"><% = frs1("Total") %>
</td>

  </tr>

<%
   
   frs1.movenext
  loop

%>



                                  </table>



                                </td>
                              </tr>
                              <tr> 
                                <td align="right" height="28"> 

                                  <%
	 if frs.recordcount>0 then
             call countpage(frs.PageCount,pageid)
			 response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
			 if Clng(pageid)<>1 then
                 response.write " <a href=javascript:gtpage('1') style='cursor:hand' >First</a> "
                 response.write " <a href=javascript:gtpage('"&(pageid-1)&"') style='cursor:hand' >Previous</a> "
			 else
                 response.write " First "
                 response.write " Previous "
			 end if
	         if Clng(pageid)<>Clng(frs.PageCount) then
                 response.write " <a href=javascript:gtpage('"&(pageid+1)&"') style='cursor:hand' >Next</a> "
                 response.write " <a href=javascript:gtpage('"&frs.PageCount&"') style='cursor:hand' >Last</a> "
             else
                 response.write " Next "
                 response.write " Last "
			 end if
	         response.write "&nbsp;&nbsp;"
	 end if
%>
                                </td>
                              </tr>
                              <tr> 
                                <td height="28" align="center"> 
<%
   response.write "<input type=hidden value='' name=whatdo>"
   response.write "<input type=hidden value="&pageid&" name=pageid>"

		  frs.close
			  set frs=nothing
			  conn.close
			  set conn=nothing
%>
                                                         
 </td>
                              </tr>
                              <tr> 
                                <td valign="top">　</td>
                              </tr>
                            </table>
                          </form>



<%
'-----------------------------------------------------------------------------
'
'      End of the main Content 
'
'-----------------------------------------------------------------------------
%>
</td>
              </tr>
                </table>
                </td>
                </tr>
              </table>  

<%

  ' function
  Sub countpage(PageCount,pageid)
  response.write pagecount&"</font> Pages "
	   if PageCount>=1 and PageCount<=10 then
		 for i=1 to PageCount
		   if (pageid-i =0) then
              response.write "<font color=green> "&i&"</font> "
		   else
             response.write " <a href=javascript:gtpage('"&i&"') style='cursor:hand' >"&i&"</a>"
		   end if
		 next
	   elseif PageCount>11 then
	      if pageid<=5 then
		     for i=1 to 10
		       if (pageid-i =0) then
                 response.write "<font color=green> "&i&"</font> "
		       else
                 response.write " <a href=javascript:gtpage('"&i&"') style='cursor:hand' >"&i&"</a>"
		       end if
		     next
		  else
		    for i=(pageid-4) to (pageid+5)
		       if (pageid-i =0) then
                 response.write "<font color=green> "&i&"</font> "
		       elseif i=<pagecount then
                 response.write " <a href=javascript:gtpage('"&i&"') style='cursor:hand' >"&i&"</a>"
		       end if
			next
		  end if
	   end if
  end sub
%>         
              </body>
              </html>