<!--#include file="include/SQLConn.inc" -->
<!--#include file ="js/OVERLIB.JS" -->
<!--#include file ="js/OVERLIB_MINI.JS" -->
<!--#include file ="js/select_date.JS" -->
<% 

' check which page is it
page_id=request("pageid")



Coupon_Type    = Request.Form("Coupon_Type")
Coupon_Batch   = Request.Form("Coupon_Batch")
Start_Range    = Request.Form("Start_Range")
End_Range      = Request.Form("End_Range")
Face_Value     = Request.Form("Face_Value")
Excel_Type     = Request.Form("Excel_Type")
Expiry_Date    = Request.Form("Expiry_Date")




%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>Shell CVS Administration</title>
<link rel="stylesheet" type="text/css" href="include/hse.css" />
<SCRIPT language=JavaScript>
<!--
function delcheck(){
k=0;
document.fm1.action="execute1.asp"
	if (document.fm1.mid!=null)
	{
		for(i=0;i<document.fm1.mid.length;i++)
		{
			if(document.fm1.mid[i].checked)
			  {
			  k=1;
			  i=1;
			  break;
			  }
		}
		if(i==0)
		{
			if (!document.fm1.mid.checked)
               k=0;
			else
               k=1;
		}
	}

if (k==0)
  alert("You must  select one record at least !");	
else if (k==1)
 {
  var msg = "Are you sure ?";
  if (confirm(msg)==true)
   {
    document.fm1.whatdo.value="delpc";
    document.fm1.submit();
   }
 }

}




function gtpage(what)
{
document.fm1.pageid.value=what;
document.fm1.action="pco_d1.asp"
document.fm1.submit();
}

function findenum()
{
document.fm1.action="pco_d1.asp"
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


 <table width="100%" border="0" cellpadding=0 cellspacing="0" bgcolor="#FFFFFF" height="100%">
                <tr>
                  <td valign="top" align="center" bgcolor="#E6EBEF">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#E6EBEF" class="normal">
                     <tr> 
                          
                        <td height="28" align="center"><font color="#FF6600"><b>
	Pro Coupon Management		</b></font></td>
                        </tr>
                        <tr> 
                          <td valign="top" align="center">
                            <form name=fm1 method=post>
                            <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" class="normal">
			 
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
      
       fsql = "select * from CouponRequest where 1 =1  "

    ' Coupon Type
   if Coupon_Type <> "" then

   fsql = fsql & " and Product_Type = '" & Coupon_Type & "' "
   
   end if

   ' Batch
   if Coupon_Batch <> "" then

   fsql = fsql & " and Batch = '" & Coupon_Batch & "' "
   
   end if

   ' Face Value
   if Face_Value <> "" then

   fsql = fsql & " and FaceValue like '%" & Face_Value & "%' "
   
   end if

   ' Check Coupon Number
   if Start_Range <> "" then

   fsql = fsql & " and Start_Range like '%" & Start_Range & "%' "
   
   end if

 ' Check Coupon Number
   if End_Range <> "" then

   fsql = fsql & " and End_Range like '%" & End_Range & "%' "
   
   end if

  ' Check Excel type
   if Excel_type <> "" then

   fsql = fsql & " and Excel_type = '" & Excel_type & "' "
   
   end if

  ' Check Expiry Date
   if Expiry_Date <> "" then

   fsql = fsql & " and Expiry_date = '" & Expiry_Date & "' "
   
   end if

  

       fsql = fsql & " order by Issue_Date desc"

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
          response.write "Total <font color=red>"&findrecord&"</font> Records ; Total <font color=blue>"

         frs.PageSize = 10

' Call the function to count the page.

         call countpage(frs.PageCount,pageid)
         end if
	     'response.write "&nbsp;&nbsp;<input type='text' name='findnum' size='13' value='"&findnum&"'>"
		 'response.write "&nbsp;&nbsp;<input type='button' value='   Search   ' onClick='findenum();' class='common'>"
	   
%>

Coupon Type
<input type="text" name="Coupon_Type" size="2" maxlength="2" value="<% = Coupon_Type %>">
Batch
<input type="text" name="Coupon_Batch" size="3" maxlength="3" value="<% = Coupon_Batch %>">
Face Value
<input type="text" name="Face_Value" size="3" maxlength="3" value="<% = Face_Value %>">
Start Range
<input type="text" name="Start_Range" size="7" maxlength="6" value="<% = Start_Range %>">
End Range
<input type="text" name="End_Range" size="7" maxlength="6" value="<% = End_Range %>">
Excel Type :
<input type="text" name="Excel_Type" size="4" value="<% = Excel_Type %>">
Expiry Date :
<input type="text" name="Expiry_Date" size="12" value="<% = Expiry_Date %>">

<input type="button" value="   Search   " onClick="findenum();" class="common">

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
<td width="5%"></td>
<td width="7%" height="28">Face Value</td>
<td >Type</td>

<td width="5%" height="28">Batch</td>

<td width="10%" height="28">Start Range</td>

<td width="10%" height="28">End Range</td>
<td >Digital</td>

<td width="10%" height="28">Category</td>
<td >Dealer</td>
<td>Canopy<br/>Copy<br/>Disc</td>
<td width="10%">Expiry Date
</td>
<td width="10%">Issue Date</td>
<td>Completed</td>
<td>Excel<br/> Type</td>
<td width="10%">Effective Date</td>
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
<td>
<input type="checkbox" name="mid" value="<% = frs("Requestid") %>">
</td>
<% id = frs("RequestID") %>
<td align=center width="95" height="28"><a href="pco2.asp?id=<% =frs("RequestID") %>"><% = frs("FaceValue")%></a></td>
<td  height="28">
<% = frs("Product_Type") %>
</td>
<td  height="28"><% = frs("Batch") %>
</td>

<td height="28"><% = frs("Start_Range") %>
</td>

<td ><% = frs("End_Range") %>
</td>

<td ><% = frs("Digital") %>
</td>

<td >
<% = frs("Category") %>
</td>

<td >
<% = frs("Dealer") %>
</td>

<td>
<% = frs("Canopy_Copy_Disc") %> 
</td>

<td>
<% = frs("Expiry_Date") %>
</td>

<td >
<% = frs("Issue_Date") %>
</td>

<td >
<% = frs("Completed") %>
</td>

<td >
<% = frs("Excel_Type") %>
</td>

<td >
<% = frs("Effective_Date") %>
</td>

</tr>
<%
   flage=not flage
   frs.movenext
  loop
 end if
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
<input type="button" value="    Delete    " onClick="delcheck();" class="common">


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