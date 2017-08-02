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
Coupon_Number  = Request.Form("Coupon_Number")




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
document.fm1.action="execute2.asp"
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
    document.fm1.whatdo.value="del_leaflet";
    document.fm1.submit();
   }
 }

}

function dosubmit(){
 document.fm1.action="Analy_NewRange1.asp"; 
 document.fm1.submit();
}

function exportCSV()
{
document.fm1.action="coupon_redempted_csv1.asp"
document.fm1.submit();
}


function gtpage(what)
{
document.fm1.pageid.value=what;
document.fm1.action="coupon_redempted1.asp"
document.fm1.submit();
}

function findenum()
{
document.fm1.action="coupon_redempted1.asp"
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
                          
                        <td height="28" align="center"><font color="#FF6600"><b>Coupon Analysis</b></font></td>
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
        findnum=replace(trim(request.form("findnum")),"%","％")
        findnum=replace(findnum,"'","''")
   
' Start the queries
      
       fsql = "select * from LeafletBatch where 1 =1  "

    ' Coupon Type
   if Coupon_Type <> "" then

   fsql = fsql & " and Coupon_Type = '" & Coupon_Type & "' "
   
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
   if Coupon_Number <> "" then

   fsql = fsql & " and Cast(Start_Range as float) <= Cast(" & Coupon_Number & " as float) "

   fsql = fsql & " and Cast(End_Range as float) >= Cast(" & Coupon_Number & " as float) "

   end if




  

       fsql = fsql & " order by ID desc"

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
Coupon Number
<input type="text" name="Coupon_Number" size="7" maxlength="6" value="<% = Coupon_Number %>">
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

<td ></td>


<td colspan="5" align="center">First Batch</td>

<td colspan="5" align="center">Second Batch</td>

<td  colspan="2"></td>

</tr>

     <tr bgcolor="#DFDFDF">

<td width="5%"></td>



<td  height="28">Face Value</td>

<td >Type</td>

<td  height="28">Batch</td>

<td height="28">Start Range</td>

<td  height="28">No. of coupon</td>
<td  height="28">Face Value</td>
<td >Type</td>

<td  height="28">Batch</td>

<td height="28">Start Range</td>

<td  height="28">No. of coupon</td>

<td  height="28">No. of leaflet</td>

<td></td>

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
<input type="checkbox" name="mid" value="<% = frs("ID") %>">
</td>



<td  height="28"><a href="coupon_redempted2.asp?id=<% =frs("ID") %>"><% = frs("FaceValue1")%></a></td>
<td  height="28">
<% = frs("Coupon_Type1") %>
</td>
<td  height="28"><% = frs("Batch1") %>
</td>

<td height="28"><% = frs("Start_Range1") %>
</td>

<td ><% = frs("NoOfCoupon1") %>
</td>


<td >
<% = frs("FaceValue2") %>
</td>

<td >
<% = frs("Coupon_Type2") %>
</td>

<td >
<% = frs("Batch2") %> 
</td>

<td height="28"><% = frs("Start_Range2") %>
</td>

<td ><% = frs("NoOfCoupon2") %>
</td>

<td ><% = frs("NoOfLeaflet") %>
</td>

<td>
<% id = frs("ID") %>
<a href="Analy_NewRange1.asp?id=<% =frs("ID") %>">Edit</a></td>

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

<input type="button" value="    New Leaflets    " onClick="dosubmit();" class="common">&nbsp;

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