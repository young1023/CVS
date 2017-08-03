<!--#include file="include/SQLConn.inc" -->
<% 

    RequestID = Request("ID")

    sql = "Select * From CouponRequest where RequestID = "&RequestID

    Set Rs = Conn.Execute(sql)

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>Shell CVS Administration</title>
<link rel="stylesheet" type="text/css" href="include/hse.css" />
<SCRIPT language=JavaScript>
<!--
function dosubmit(){
 document.fm1.action="execute1.asp?id=<%=RequestID%>";
 
 if (document.fm1.Face_Value.value == "")
  {
   alert("Please input face value.");
   document.fm1.Face_Value.focus();
   return false;
  }
if (document.fm1.Product_Type.value == "")
  {
   alert("Please input coupon type.");
   document.fm1.Product_Type.focus();
   return false;
  }

document.fm1.submit();
}




//-->
</SCRIPT>
</head>

<body leftmargin="0" topmargin="0" >
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


<form name="fm1" method="post" action="">

  <table width="99%" border="0" class="normal">
    <tr> 
      <td colspan="2" class="BlueClr"><font size="5" face="Times New Roman, Times, serif">Edit Pro Coupon </font></td>
    </tr>
    <tr> 
      <td width="35%"> </td>
      <td width="65%" align="right"></td>
    </tr>
    <tr> 
      <td colspan="2"></td>
    </tr>
  <tr> 
      <td colspan="2"  align="right">
<font color="red">*</font>Denotes a mandatory field</td>
    </tr>
    <tr> 
      <td colspan="2" bordercolor="#999966" class="ValidHead"></td>
    </tr>

 <tr><td>Face Value: </td>
      <td>
<input name="Face_Value" type=text value="<% = rs("FaceValue") %>">	    
</td>
   </tr>
 <tr>
<td>Type:</td>
<td>
<input name="Product_Type" type=text value="<% = rs("Product_Type") %>">	    
</td>
</tr>
     
 <tr><td>Batch: </td>
      <td>
<input name="Batch" type=text value="<% = rs("Batch") %>">	    
</td>
   </tr>

<tr><td>Start Range:</td>
      <td>
<input name="Start_Range" type=text value="<% = rs("Start_Range") %>">	    
</td>
   </tr>
<tr><td>End Range:</td>
      <td>
<input name="End_Range" type=text value="<% = rs("End_Range") %>">	    
</td>
   </tr>

     <tr>
	<td><b>Coupon Value</b></td>
      <td>

</td>
   	</tr>

<tr><td>Digital Coupon:</td>
      <td>
<select size="1" name="Digital" class="common">
			<option value="N" <%if Trim(rs("Digital"))="N" Then %>Selected<%End If%>>No</option>
			<option value="Y" <%if Trim(rs("Digital"))="Y" Then %>Selected<%End If%>>Yes</option>

	</select>    
</td>
   </tr>

<tr><td>Category:</td>
      <td>
<input name="Category" type=text value="<% = rs("Category") %>">	    
</td>
   </tr>
<tr><td>Dealer A/C:</td>
      <td>
<input name="Dealer_AC" type=text value="<% = rs("Dealer") %>">	    
</td>
   </tr>
<tr><td>Canopy Disc:</td>
      <td>
<input name="Canopy_Copy_Disc" type=text value="<% = rs("Canopy_Copy_Disc") %>">	    
</td>
   </tr>
<tr><td>Expiry Date: (mm-dd-yyyy)</td>
      <td>
<input name="Expiry_Date" type=text value="<% = rs("Expiry_Date") %>">	    
</td>
   </tr>
<tr><td>Issue Date: (mm-dd-yyyy)</td>
      <td>
<input name="Issue_date" type=text value="<% = rs("Issue_Date") %>">	    
</td>
   </tr>
<tr><td>Completed:</td>
      <td>
<input name="Completed" type=text value="<% = rs("Completed") %>">	    
</td>
   </tr>
<tr><td>Excel Type:</td>
      <td>
<input name="Excel_Type" type=text value="<% = rs("Excel_Type") %>">	    
</td>
   </tr>
<tr><td>Effective Date:</td>
      <td>
<input name="Effective_Date" type=text value="<% = rs("Effective_Date") %>">	    
</td>
   </tr>

<tr><td>Time Allowed:</td>
      <td>From:&nbsp;



<select size="1" name="Start_Time" class="common">

<% 
    For i = 0 to 24

%>
			<option value="<% = i %>" <% If trim(i)  = trim(rs("Start_Time")) then %>Selected<% End If %>><% = i %></option>
		
<%
    Next

%>

	</select>&nbsp;To:&nbsp;

<select size="1" name="Start_Time" class="common">

<% 
    
    For j = 0 to 24

%>
			<option value="<% = j %>" <% If trim(j)  = Trim(rs("End_Time")) then %>Selected<% End If %>><% = j %></option>
		
<%
    Next

%>

	</select>    
</td>
   </tr>

<tr> 
      <td>Restricted Station:  </td>
<td >
<%

		SQL1 = "Select * from Station where Cast(station as int) >= 200 Order by station"

		Set Rs1 = Conn.Execute(SQL1)

        SQL2 = "Select Station From Restricted_Station where RequestID = "& RequestID

        'response.write sql2

		Set Rs2 = Conn.Execute(SQL2)


        Rs1.MoveFirst
						
            If Rs1.EoF Then

                  Response.write "No station found."

						Else
					        
                 i = 1
							
                   Do While Not Rs1.EoF


                          

 %>

<Input Type="checkbox" Name="Station" Value="<% = Rs1("Station") %>">
 
   <% = Rs1("Station") %>
                        
<%    
   
     Rs1.Movenext
                               
       If Len(i/ 10) = 1 then
                               
           Response.write "<br>"
 
                               End If

                                  i = i + 1

							Loop

						End If
%>
        
</td>
    </tr>
	<tr>
<td></td>
<td>
  
</td>
	</tr>
     <tr>
 
    
                                        <td colspan="2" align="center"> 

<input type="button" value="  Submit  " onClick="dosubmit();" class="common">
<input type=hidden name=whatdo value="edit_pco">

</td>
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