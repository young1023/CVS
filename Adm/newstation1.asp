<!--#include file="include/SQLConn.inc" -->
<%
   

    IPAddress = Request("IPAddress")


    If IPAddress <> "" then

      sql = "Select * from Station where IPAddress = '"& IPAddress &"'"

      Set rs = Conn.Execute(sql)


    Station = rs("Station")

    ShipToCode = rs("ShipToCode")

    SoldToCode = rs("SoldToCode")

    SAPCode = rs("SAPCode")

    Outdoor = trim(rs("Outdoor"))

    End if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>Shell CVS Administration</title>
<link rel="stylesheet" type="text/css" href="include/hse.css" />
<SCRIPT language=JavaScript>
<!--
function dosubmit(){
 document.fm1.action="execute1.asp";
 
 if (document.fm1.Station.value == "")
  {
   alert("Please input station.");
   document.fm1.Station.focus();
   return false;
  }
if (document.fm1.IPAddress.value == "")
  {
   alert("Please input IP Address.");
   document.fm1.IPAddress.focus();
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
      <td colspan="2" class="BlueClr"><font size="5" face="Times New Roman, Times, serif">Create Station </font></td>
    </tr>
  <tr> 
      <td colspan="2"  align="right">
<font color="red">*</font>Denotes a mandatory field</td>
    </tr>
    <tr> 
      <td colspan="2" bordercolor="#999966" class="ValidHead"></td>
    </tr>

 <tr><td>Station: </td>
      <td>
<input name="Station" type=text value="<% = station %>">	    
</td>
   </tr>
 <tr>
<td>IP Addres:</td>
<td>
<input name="IPAddress" type=text value="<% =IPAddress %>">	    
</td>
</tr>
     
 <tr><td>ShipToCode: </td>
      <td>
<input name="ShipToCode" type=text value="<% =ShipToCode %>">	    
</td>
   </tr>

<tr><td>Sold To Code:</td>
      <td>
<input name="SoldToCode" type=text value="<% =SoldToCode %>">	    
</td>
   </tr>
<tr><td>SAP Code:</td>
      <td>
<input name="SAPCode" type=text value="<% =SAPCode %>">	    
</td>
   </tr>

<tr><td>Outdoor:</td>
      <td>
	<select size="1" name="Outdoor" class="common">
			<option value="0" <%if Outdoor = 0 then%>Selected<%End If%>>No</option>
			<option value="1" <%if Outdoor = 1 then%>Selected<%End If%>>Yes</option>

	</select>    
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
<input type="Reset" value="  Reset  " class="common">

<%  If IPAddress <> "" Then %>
<input type=hidden name=whatdo value='edit_st'>
<% Else %>
<input type=hidden name=whatdo value='add_st'>
<%  end if %>

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