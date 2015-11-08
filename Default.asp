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

Alert = "油站編號未設置, 請提供此電腦的 IP 地址 [ " & UserIPAddress & " ] 給系統管理員"


End If

%>

<HTML>
<HEAD>
<title>
    --  禮券驗證系統  -- 
</title>
<meta http-equiv="Content-Language" content="zh-hk">
<meta http-equiv="Content-Type" content="text/html; charset=big5" />
<link rel="stylesheet" type="text/css" media="all" href="include/publish.css" />
<link rel="shortcut icon" href="images/favicon.ico" />
<!--#include file="include/JavaScript.js" -->
</HEAD>
<body class="homepage" onload="document.fm1.UserID.focus();">
<!--#include file="include/header.inc" -->

<form name="fm1" method="post" action="">

<div align="center">

<h1 class="Title">禮券驗證系統</h1>

<table width="60%" border="0" class="Login">

    <tr>

	    <td width="40%">油站編號 

        </td> 

		<td >
					<% = StationID %>
        
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
 <tr>

	    <td width="40%">

        </td> 

		<td ><p>This is the testing server of CVS system.</p>
             <p>All the data are for testing purpose only.</p>

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

</body>
</HTML>
