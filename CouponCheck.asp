<!--#include file="include/SQLConn.inc" -->

<%

UserIPAddress = Request.ServerVariables("REMOTE_ADDR")

response.write UserIPAddress



%>

<HTML>
<HEAD>
<title>
 - §�����Ҩt�� (�D��) -
</title>
<meta http-equiv="Content-Language" content="zh-hk">
<meta http-equiv="Content-Type" content="text/html; charset=big5" />
<link rel="stylesheet" type="text/css" media="all" href="include/publish.css" />
<link rel="shortcut icon" href="images/favicon.ico" />

<!--#include file="include/JavaScript.js" -->
</HEAD>
<body class="homepage" onload="document.barCodeForm.Barcode.focus();">

<form name="barCodeForm" method="post" action="" onSubmit="javascript:return mod10CheckDigit();">

<div align="center">

<h1 class="Title">§�����Ҩt�� (�D��)</h1>


<table width="60%" border="0" cellpadding="04" cellspacing="0" class="Report">

      <tr>

              <td>

                      <a href="DailyReport.asp">���Ҭ���</a>                 

              </td>

               <td>

                      <a href="logout.asp">�n�X</a>                 

              </td>

      </tr>

</table>

<br/>



<table width="60%" border="0"  cellpadding="04" cellspacing="2" class="Login">

        

    <tr>

	    <td width="23%">§�鸹�X 

        </td> 

		<td width="37%">

					
<input name="Barcode" type="text" autocomplete="off" value="" size="20" maxlength="15" id="bc" ">


	<input type="Submit" Name="Button" value="�T�w">


        </td>

     </tr>


</table>

</div>

</form>

<div id="myDiv"></div>

</body>
</HTML>
