<!--#include file="include/SQLConn.inc" -->

<!--#include file ="js/OVERLIB.JS" -->

<!--#include file ="js/OVERLIB_MINI.JS" -->

<!--#include file ="js/select_date.JS" -->

<% 


' check which page is it


pageid=request("pageid")


From_Date      = Request.Form("From_Date")
if From_Date = "" then
   From_Date =  year(now()) & "-" & month(now()) & "-" & day(now()) 
end if

To_Date        = Request.Form("To_Date")

if To_Date = "" then
   To_Date = year(now()) & "-" & month(now()) & "-" & day(now())
end if

Coupon_Type    = Request.Form("Coupon_Type")
Coupon_Batch   = Request.Form("Coupon_Batch")
Start_Range    = Request.Form("Start_Range")
End_Range      = Request.Form("End_Range")
Face_Value     = Request.Form("Face_Value")



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
document.fm1.action="rd_co2.asp"
document.fm1.submit();
}



function dateCheck(inputText) {

         debugger;

         var dateFormat = /^(0?[1-9]|1[012])[\/\-](0?[1-9]|[12][0-9]|3[01])[\/\-]\d{4}$/;

          var flag = 1;


         if (inputText.value.match(dateFormat)) {

           var inputFormat1 = inputText.value.split('/');

             var inputFormat2 = inputText.value.split('-');

             linputFormat1 = inputFormat1.length;

             linputFormat2 = inputFormat2.length;

 

             if (linputFormat1 > 1) {

                 var pdate = inputText.value.split('/');

             }

             else if (linputFormat2 > 1) {

                 var pdate = inputText.value.split('-');

             }

             var date = parseInt(pdate[0]);

             var month = parseInt(pdate[1]);

             var year = parseInt(pdate[2]);

 

             var ListofDays = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];

             if (month == 1 || month > 2) {

                 if (date > ListofDays[month - 1]) {

                     alert("Invalid date format!");

                     return false;

                 }

             }

 

             if (month == 2) {

                 var leapYear = false;

 

                 if ((!(year % 4) && year % 100) || !(year % 400)) {

                     leapYear = true;

 

                 }

                 if ((leapYear == false) && (date >= 29)) {

                     alert("Invalid date format!");

                     return false;

                 }

                 if ((leapYear == true) && (date > 29)) {

                     alert("Invalid date format!");

                     return false;

                 }

             }

         }

         else {

             alert("Invalid date format!");

             return false;

         }

     }





function findenum()
{
document.fm1.action="rd_co2.asp"
document.fm1.submit();
}

function exportExcel()
{
document.fm1.action="rd_co_excel2.asp"    
document.fm1.submit();
}


function dateCheck(inputText) {
         debugger;

         var dateFormat = /^(0?[1-9]|[12][0-9]|3[01])[\/\-](0?[1-9]|1[012])[\/\-]\d{4}$/;

         var flag = 1;

         if (inputText.value.match(dateFormat)) {
             document.myForm.dateInput.focus();

             var inputFormat1 = inputText.value.split('/');
             var inputFormat2 = inputText.value.split('-');
             linputFormat1 = inputFormat1.length;
             linputFormat2 = inputFormat2.length;

             if (linputFormat1 > 1) {
                 var pdate = inputText.value.split('/');
             }
             else if (linputFormat2 > 1) {
                 var pdate = inputText.value.split('-');
             }
             var date = parseInt(pdate[0]);
             var month = parseInt(pdate[1]);
             var year = parseInt(pdate[2]);

             var ListofDays = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
             if (month == 1 || month > 2) {
                 if (date > ListofDays[month - 1]) {
                     alert("Invalid date format!");
                     return false;
                 }
             }

             if (month == 2) {
                 var leapYear = false;

                 if ((!(year % 4) && year % 100) || !(year % 400)) {
                     leapYear = true;

                 }
                 if ((leapYear == false) && (date >= 29)) {
                     alert("Invalid date format!");
                     return false;
                 }
                 if ((leapYear == true) && (date > 29)) {
                     alert("Invalid date format!");
                     return false;
                 }
             }
             if (flag == 1) {
                 alert("Valid Date");
             }
         }
         else {
             alert("Invalid date format!");
             document.myForm.dateInput.focus();
             return false;
         }
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
Redemption Coupon</b></font></td>
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
        findnum=replace(trim(request.form("findnum")),"%","�H")
        
findnum=replace(findnum,"'","''")
   
' Start the queries
         
      

      Search_No =  pageid * 10 + 1000

      fsql = "SELECT  Top " & Search_No

      fsql = fsql & " Convert(datetime, Present_Date,103) as [Present Date] , m.RequestedID as Station, "

      'fsql = fsql & " m.Coupon_Type as [Coupon Type], m.Coupon_Batch as [Batch],"

      fsql = fsql & " m.Coupon_Number as [Coupon Number], "

      'fsql = fsql & " m.SaleAmount as [Sale Amount], "

      fsql = fsql & " Convert(varchar, c.Expiry_Date,103) as [Expiry Date] "

      fsql = fsql & " From MasterCoupon m Left Join CouponRequest c on m.Coupon_Type = c.Product_Type and "

      
fsql = fsql & " m.Coupon_batch = c.Batch and Cast(m.Face_Value as float) = Cast(c.FaceValue as float) and "

      fsql = fsql & " c.Start_Range <= m.Coupon_Number and c.End_Range >= m.Coupon_Number where "

      fsql = fsql & " (Datediff(day, Present_Date, '"&From_Date&"') < = 0 or '"&From_Date&"' = '') "

      fsql = fsql & " and (Datediff(day, Present_Date, '"&To_Date&"') >= 0 or '"&To_Date&"' = '') "

      fsql = fsql & " and (m.RequestedID = '"&Station&"'  or '"&Station&"' = '') and "

      fsql = fsql & " (m.Coupon_Type = '"&Coupon_Type&"'  or '"&Coupon_Type&"' = '') and "

      fsql = fsql & " (m.Coupon_Batch = '"&Coupon_Batch&"' or '"&Coupon_Batch&"' = '') and "

      fsql = fsql & " (Cast(m.Face_Value as float)= '"&Face_Value&"' or '"&Face_Value&"' = '') and "

      fsql = fsql & " (m.Coupon_Number >= '"&Start_Range&"' or '"&Start_Range&"' = '') and "

      fsql = fsql & " (m.Coupon_Number <= '"&End_Range&"' or '"&End_Range&"' = '') "

      fsql = fsql & " group by Convert(datetime, Present_Date,103), m.requestedid, "

      'fsql = fsql & " m.Coupon_type, m.Coupon_Batch, "

      
fsql = fsql & " m.Coupon_Number, "

      'fsql = fsql & " m.SaleAmount, 

      fsql = fsql & " c.Expiry_Date "

      fsql = fsql & "  Order by Convert(datetime, Present_Date,103), m.Requestedid, "

      fsql = fsql & " m.Coupon_Number, c.Expiry_date"

    
      

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
          
'response.write "Total <font color=red>"&findrecord&"</font> Records ; Total <font color=blue>"

         
frs.PageSize = 10
       end if
%>


Date From:
<input type="text" name="From_Date" size="10" value="<% = From_Date %>" onkeyup="dateCheck(document.fm1.From_Date);">
<a href="javascript:show_calendar('fm1.From_Date');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a full year pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;"><img src="images/show-calendar.gif" width=24 height=22 border=0></a>
To Date:
<input type="text" name="To_Date" size="10" value="<% = To_Date %>" onkeyup="dateCheck(document.fm1.To_Date);">
<a href="javascript:show_calendar('fm1.To_Date');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a full year pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;"><img src="images/show-calendar.gif" width=24 height=22 border=0></a>
Coupon Type
<input type="text" name="Coupon_Type" size="2" maxlength="2" value="<% = Coupon_Type %>">
Batch
<input type="text" name="Coupon_Batch" size="3" maxlength="3" value="<% = Coupon_Batch %>">
Face Value
<input type="text" name="Face_Value" size="3" maxlength="3" value="<% = Face_Value %>">
Start Range
<input type="text" name="Start_Range" size="6" maxlength="6" value="<% = Start_Range %>">
End Range
<input type="text" name="End_Range" size="6" maxlength="6" value="<% = End_Range %>">

<input type="button" value="   Search   " onClick="findenum();" class="common">

<% if From_Date <> "" Then %>

<input type="button" value="   Export to Excel   " onClick="exportExcel();" class="common">

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

<td height="28">Present Date</td>
<td height="28">Station</td>

<td height="28">Coupon Number</td>
<td height="28">Expiry Date</td>

</tr>
 
<%
 i=0
 if frs.recordcount>0 then
  frs.AbsolutePage = pageid
  do while (frs.PageSize-i)
   if frs.eof then exit do
   i=i+1
   if flage then
     mycolor="#ffffff"
   else
	 mycolor="#efefef"
   end if
  
%>
   <tr>

<td align=center width="95" height="28"><% = frs("Present Date")%></td>
<td  height="28"><% = frs("Station") %>
</td>

<td height="28"><% = frs("Coupon Number") %>
</td>

<td height="28"><% = frs("Expiry Date") %>
</td>
</tr>
<%
   
   frs.movenext
  loop

End if
 
  %>

                                  </table>
                                </td>
                              </tr>
                              
<tr> 
                                
<td align="right" height="28"> 

                                  
<%
	 if frs.recordcount>0 then
             'call countpage(frs.PageCount,pageid)
			 
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
                                <td valign="top"></td>
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