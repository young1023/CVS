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

Alert = "�o���s�����]�m, �д��Ѧ��q���� IP �a�} [ " & UserIPAddress & " ] ���t�κ޲z��"


End If

%>

<HTML>
<HEAD>
<title>
    --  §�����Ҩt��  -- 
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

<h1 class="Title">§�����Ҩt��</h1>

<table width="60%" border="0" class="Login">

    <tr>

	    <td width="40%">�o���s�� 

        </td> 

		<td >
					<% = StationID %>
        
        </td>

     </tr>

     <tr>

	    <td width="40%">�Τ�s��

        </td> 

		<td >

			<input name="UserID" type=text value="" size="20" <%If Alert<>"" Then%>Readonly<%End If%> autocomplete="off">

        </td>

     </tr>

     <tr>

	    <td width="40%">�Τ�K�X

        </td> 

		<td >

			<input name="Password" type=password value="" size="20" autocomplete="off" <%If Alert<>"" Then%>Readonly<%End If%>>

        </td>

     </tr>

     <tr>

	    <td width="40%">

        </td> 

		<td ><input Type="Button" Name=" �T�w " value="   �T�w   " onClick="doLogin();">
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

           Response.write "�Τ�s���αK�X�����T"

        End If

        response.write Alert

  %>

</div>

</body>
</HTML>
