<!--#include file="include/SQLConn.inc" -->
<%
   

    UserName = Request("Username")


    If Username <> "" then

      sql = "Select * from CVSUser where UserName = '"& Username &"'"

      Set rs = Conn.Execute(sql)


    UserName = rs("UserName")

    Password = rs("Password")

    SecLevel = rs("SecLevel")

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
 
 if (document.fm1.UserName.value == "")
  {
   alert("Please input User Name.");
   document.fm1.UserName.focus();
   return false;
  }
if (document.fm1.Password.value == "")
  {
   alert("Please input Password.");
   document.fm1.Password.focus();
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

 <tr><td>User Name: </td>
      <td>
<input name="UserName" type=text value="<% = trim(UserName) %>">	    
</td>
   </tr>
 <tr>
<td>Password:</td>
<td>
<input name="Password" type=text value="<% = trim(Password) %>">	    
</td>
</tr>

<tr><td>User Level:</td>
      <td>
	<select size="1" name="SecLevel" class="common">
			<option value="7" <%if SecLevel = 7 then%>Selected<%End If%>>Administrator</option>
			<option value="5" <%if SecLevel = 5 then%>Selected<%End If%>>Reporter</option>
			<option value="3" <%if SecLevel = 3 then%>Selected<%End If%>>Operator</option>
			<option value="1" <%if SecLevel = 1 then%>Selected<%End If%>>Station User</option>

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

<%  If UserName <> "" Then %>
<input type=hidden name=whatdo value='edit_usr'>
<% Else %>
<input type=hidden name=whatdo value='add_usr'>
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