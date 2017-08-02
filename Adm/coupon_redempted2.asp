<!--#include file="include/SQLConn.inc" -->
<!--#include file ="js/OVERLIB.JS" -->
<!--#include file ="js/OVERLIB_MINI.JS" -->
<!--#include file ="js/select_date.JS" -->
<% 

' check which page is it
page_id=request("pageid")

ID = Request("ID")

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
    document.fm1.whatdo.value="del_new_range";
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
document.fm1.action="coupon_redempted_csv1.asp?id=<% =id %>"
document.fm1.submit();
}


function gtpage(what)
{
document.fm1.pageid.value=what;
document.fm1.action="coupon_redempted2.asp?id=<% =id %>"
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

       'Update redeem at coupon redemption table
       Set rs = server.createobject("adodb.recordset")

	   rs.open ("Exec Update_Redemption_Rate '"& ID &"'") ,  conn,3,1

		  pageid=1

		end if

   
' Start the queries
      
       fsql = "Select a.LeafLetNo as 'LeafLetNo' , a.FaceValue as 'FaceValue1', "

       fsql = fsql & "a.Coupon_Type as 'Coupon_Type1', a.Batch as 'Batch1', "

       fsql = fsql & "a.Start_Range as 'Start_Range1', a.End_Range as 'End_Range1', "

       fsql = fsql & " a.NoOfCoupon as 'TotalNo1', b.FaceValue as 'FaceValue2', "

       fsql = fsql & "b.Coupon_Type as 'Coupon_Type2', b.Batch as 'Batch2', "

       fsql = fsql & "b.Start_Range as 'Start_Range2', b.End_Range as 'End_Range2', "

       fsql = fsql & "b.NoOfCoupon as 'TotalNo2', TotalNo = (a.NoOfCoupon + b.NoOfCoupon ), "

       fsql = fsql & "TotalRedeem = (a.Redeem + b.Redeem) "


       fsql = fsql & "from Coupon_Redemption a inner join coupon_Redemption b "

       fsql = fsql & "on a.leafletid = b.leafletid and a.leafletno = b.leafletno "

       fsql = fsql & "and a.id = (b.id - 1) where a.leafletid =" & ID

       fsql = fsql & " order by a.LeafLetID "


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
	   
%>


   </td>
      </tr>
         <tr> 
            <td valign="top" height="28">


   <table border="0" align=center cellpadding="1" width="100%" cellspacing="1" class="normal">

     <tr bgcolor="#DFDFDF">

<td ></td>


<td colspan="5" align="center">First Batch</td>

<td colspan="5" align="center">Second Batch</td>

<td  colspan="3"></td>

</tr>

     <tr bgcolor="#DFDFDF">

<td  width="5%" height="28">#</td>
<td  height="28">Face Value</td>
<td >Type</td>

<td  height="28">Batch</td>

<td height="28">Start Range</td>

<td  height="28">End Range</td>

<td  height="28">Face Value</td>
<td >Type</td>

<td  height="28">Batch</td>

<td height="28">Start Range</td>

<td  height="28">End Range</td>

<td>Coupons Used</td>
<td>Coupons Issued</td>
<td width="10%">Redemption Rate</td>

</tr>
                                    <%



 i=0
 if frs.recordcount>0 then

  frs.AbsolutePage = pageid

  do while (frs.PageSize-i)

   if frs.eof then exit do

   i=i+1

   

  b = i mod 2

  If b <> 0 then


     mycolor="#ffffff"

   else

	 mycolor="#efefef"

   end if

  
%>
   <tr bgcolor="<% = mycolor %>">



<% LeafLetNo = frs("LeafLetNo") %>


<td>

<% = LeafLetNo %>

</td>

<td align=center width="95" height="28"><% = frs("FaceValue1")%></td>

<td  height="28">

<% = frs("Coupon_Type1") %>

</td>

<td  height="28"><% = frs("Batch1") %>
</td>

<td height="28"><% = frs("Start_Range1") %>
</td>

<td ><% = frs("End_Range1") %>
</td>

<td align=center width="95" height="28"><% = frs("FaceValue2")%></td>

<td  height="28">

<% = frs("Coupon_Type2") %>

</td>

<td  height="28"><% = frs("Batch2") %>
</td>

<td height="28"><% = frs("Start_Range2") %>
</td>

<td ><% = frs("End_Range2") %>
</td>


<td >
<% 
    = frs("TotalRedeem") 

%>
</td>

<td >

<% = frs("TotalNo") %>
</td>

<td >
  <% 


   response.write FormatNumber(frs("TotalRedeem")/frs("TotalNo") * 100,1) & "%"

   %> 
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
<input type="button" value="    csv    " onClick="exportCSV();" class="common">&nbsp;



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