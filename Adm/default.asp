<!--#include file="include/SQLConn.inc" -->
<%

Message = Request("Message")

Dim UserIPAddress
UserIPAddress = Request.ServerVariables("REMOTE_ADDR")
'response.write UserIPAddress



Set Rs1 = Server.CreateObject("Adodb.Recordset")
  
sql1 = "Select Station From Station Where IPAddress ='" &UserIPAddress& "'"

Set Rs1 = Conn.Execute(sql1)

If Not Rs1.EoF Then

StationID = Rs1("Station")

Else 

Alert = "編號未設置, 請提供此電腦的 IP 地址 [ " & UserIPAddress & " ] 給系統管理員"


End If

%>
<html>
<head>
<title>CVS Admminstration</title>
<meta http-equiv="Content-Language" content="zh-hk">
<meta http-equiv="Content-Type" content="text/html; charset=big5" />
<link rel="stylesheet" type="text/css" media="all" href="include/hse.css" />
<link rel="shortcut icon" href="images/favicon.ico" />
<!--#include file="include/JavaScript.js" -->
</head>
<html>

<body leftmargin="0" topmargin="0" onload="document.fm1.UserID.focus();">

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><img src="images/Shell.gif" /></td>
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
          <colgroup>
            <col width="15" /><col align="middle" />
            </colgroup>
            <tr>
              <td></td>
              <td><font class="Head" style="font-size: 13px">CVS Administration
              </font></td>
            </tr>
            </table>
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
      <td width="180" valign="top">
      <table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0" class="Common">
        <tr valign="top" align="center">
          <td class="HSEBlue" height="21"></td>
        </tr>
        <tr valign="top" align="center">
          <td>
          </td>
        </tr>
      </table>
      </td>
      <td width="0" class="HSEBlue"></td>
      <td valign="top">
      <table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="4" class="HSEBlue"></td>
        </tr>
        <tr valign="top">
          <td height="25"><img src="images/Curve.gif" width="22" height="16" /></td>
        </tr>
        <tr valign="top">
          <td height="100%" align="middle" >


<form name="fm1" method="post" action="">

<div align="center">



		
<% 

If  StationID = 101 then

    Alert = "No unauthority PC is allowed to access the system."

End If



%>

<table width="60%" border="0" Height="400"  class="Login">

    <tr>

	    <td width="40%">

        </td> 

		<td >

        
        </td>

     </tr>

     <tr>

	    <td width="40%">用戶編號

        </td> 

		<td >

			<input name="UserID" type=text value="" size="20" <%If Alert<>"" Then%>Readonly<%End If%> autocomplete="off">

        </td>

     </tr>

     <tr>

	    <td width="40%">用戶密碼

        </td> 

		<td >

			<input name="Password" type=password value="" size="20" autocomplete="off" <%If Alert<>"" Then%>Readonly<%End If%>>

        </td>

     </tr>

     <tr>

	    <td width="40%">

        </td> 

		<td ><input Type="Button" Name=" 確定 " value="   確定   " onClick="doLogin();">
             <input name="StationID" type="hidden" value="<%=StationID%>">


        </td>

     </tr>
 <tr>

	    <td width="40%">

        </td> 

		<td >&nbsp;


        </td>

     </tr>

    


</table>

</div>

</form>



<div class="Message">

 <%

        If Message = "Fail" then

           Response.write "用戶編號或密碼不正確"

        End If

        response.write Alert

  %>

</div>



</td>
              </tr>
                </table>
                </td>
                </tr>
              </table>
              </body>
           </html>