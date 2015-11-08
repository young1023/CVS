<!--#include file="include/SQLConn.inc" -->
<% 

' check which page is it
page_id=request("page_id")

' get the today date
submit_date = datevalue(date())





%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>Shell CVS Administration</title>
<link rel="stylesheet" type="text/css" href="include/hse.css" />
<SCRIPT language=JavaScript>
<!--
function dosubmit(){
 document.fm1.action="execute.asp";
 
 if (document.fm1.Face_Value.value == "")
  {
   alert("Please input face value.");
   document.fm1.Face_Value.focus();
   return false;
  }
if (document.fm1.Coupon_Type.value == "")
  {
   alert("Please input coupon type.");
   document.fm1.Coupon_Type.focus();
   return false;
  }

document.fm1.submit();
}

function formHandler(fm1){
var URL = document.fm1.treeid.options[document.fm1.treeid.selectedIndex].value;
URL = 'lstlocation.asp?treeid='  + URL
window.location.href = URL;
}

function doapprove(){
 
document.fm1.action="hsemis.asp?page_id=execute";
document.fm1.whatdo.value="approve_cm";
document.fm1.submit();
}


function dosave(){
document.fm1.action="hsemis.asp?page_id=execute";
  if (document.fm1.PCDate.value == "")
  {
   alert("Please enter the date.");
   document.fm1.PCDate.focus();
   return false;
  }



//-->
</SCRIPT>
</head>

<body leftmargin="0" topmargin="0" >
<div id="overDiv" style="position:absolute; visibility:hidden; z-index:1000;"></div>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="60"><IMG style="LEFT: 0px; POSITION: absolute; TOP: 0px"  alt ="Shell logo" src="IMAGES/Shell.gif" border=0 ></IMG></td>
    <td align="right"></td>
  </tr>
</table>
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
      <td width="120" valign="top">
      <table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0" class="Common">
        <tr valign="top" align="center">
          <td class="HSEBlue" height="21"></td>
        </tr>
        <tr valign="top" align="center">
          <td>
              <br/>
              <p>
              <a href="pco.asp" title="Create new pro coupon" style="TEXT-DECORATION: none">
              New Pro Coupon<br>
              </a>
              </p>
            
              <p>
              <a href="pco_d.asp" title="pro coupon" style="TEXT-DECORATION: none">
              Pro Coupon<br>
              </a>
              </p>

 <p>
              <a href="master.asp" title="pro coupon" style="TEXT-DECORATION: none">
              Master List<br>
              </a>
              </p>

 <p>
              <a href="station.asp" title="pro coupon" style="TEXT-DECORATION: none">
              Station<br>
              </a>
              </p>

 <p>
              <a href="price.asp" title="pro coupon" style="TEXT-DECORATION: none">
              Price<br>
              </a>
              </p>
          </td>
        </tr>
      </table>
      </td>
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
	Master Coupon</b></font></td>
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
        findnum=replace(trim(request.form("findnum")),"%","％")
        findnum=replace(findnum,"'","''")
   
' Start the queries
      
       fsql = "select * from MasterCoupon where 1 =1  "

  ' Search location
  ' ****************
   if location_id <> "" then
       fsql = fsql & " and Customer_Name = '"&System_Name&"' "
   end if

  ' Search Job Status
  ' *******************
   if S1 <> "" then
       fsql = fsql & " and jobstatus like '%"&S1&"%'"
   end if 

  ' Search Assign To
  ' *****************
   if assignto <> "" then
       fsql = fsql & " and AssignedTo = '"&assignto&"'"
   end if

  ' Search by Date
  ' **************
   if Sdate <> "" then
      fsql = fsql & " and TargetCompletionDate >= #"& SDate &"# and TargetCompletionDate < #"& NDate &"# "   
   end if

  ' Searh by Order Number
  ' ***********************
   if findnum <> "" then
      fsql = fsql & " and OrderNo like '%"&findnum&"%'"
   end if
  

       fsql = fsql & " order by ID desc"

        response.write fsql
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

        ' call countpage(frs.PageCount,pageid)
         end if
	     response.write "&nbsp;&nbsp;<input type='text' name='findnum' size='13' value='"&findnum&"'>"
		 response.write "&nbsp;&nbsp;<input type='button' value='   Search   ' onClick='findenum();' class='common'>"
	   
%>
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
<td width="26"></td>
<td width="95" height="28">Present Date</td>
<td height="28">Station</td>

<td width="100" height="28">Coupon<br/>Type</td>

<td width="100" height="28">Batch</td>

<td width="100" height="28">Coupon Number</td>

<td width="139" height="28">Product<br/>Type</td>
<td width="108">Car Plate no.</td>
<td>Canopy<br/>Face<br/>Value</td>
<td width="58">Expiry Date
</td>
<td>Issue Date</td>
<td>Completed</td>
<td>Excel<br/> Type</td>
<td>Effective Date</td>
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
<td width="26">
<input type="checkbox" name="mid" value="<% = frs("id") %>">
</td>

<td align=center width="95" height="28"><% = frs("Present_Date")%></td>
<td  height="28">
<% = frs("RequestedID") %>
</td>
<td  height="28"><% = frs("Coupon_Type") %>
</td>

<td height="28"><% = frs("Coupon_Batch") %>
</td>

<td  height="28"><% = frs("Coupon_number") %>
</td>

<td width="139" height="28">
<% = frs("Coupon_Type") %>
</td>

<td width="108">
<% = frs("Car_ID") %>
</td>

<td>
<% '= frs("Canopy_Copy_Disc") %> 
</td>

<td width="58">
<% '= frs("Expiry_Date") %>
</td>

<td width="58">
<% '= frs("Issue_Date") %>
</td>

<td width="58">
<% '= frs("Completed") %>
</td>

<td width="58">
<% '= frs("Excel_Type") %>
</td>

<td width="58">
<% '= frs("Effective_Date") %>
</td>

<td width="34">
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
<script language=JavaScript>
<!--
function delcheck(){
k=0;
document.fm1.action=""execute.asp"
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
  alert("You must select at least one record!");	
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
document.fm1.action="pco_d.asp"
document.fm1.submit();
}

function findenum()
{
document.fm1.action="pco_d.asp"
document.fm1.submit();
}
//-->
</script>
                                  <%
	 if frs.recordcount>0 then
            ' call countpage(frs.PageCount,pageid)
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
   response.write "<input type='button' value='    Delete    ' onClick='delcheck();' class='common'>"
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
           
              </body>

              </html>