<!--#include file="include/SQLConn.inc" -->

<%

UserIPAddress = Request.ServerVariables("REMOTE_ADDR")

response.write UserIPAddress



%>

<HTML>
<HEAD>
<title>
 - 禮券驗證系統 (澳門) -
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

<h1 class="Title">禮券驗證系統 (澳門)</h1>


<table width="60%" border="0" cellpadding="04" cellspacing="0" class="Report">

      <tr>

              <td>

                      <a href="DailyReport.asp">驗證紀錄</a>                 

              </td>

               <td>

                      <a href="logout.asp">登出</a>                 

              </td>

      </tr>

</table>

<br/>



<table width="60%" border="0"  cellpadding="04" cellspacing="2" class="Login">

        

    <tr>

	    <td width="23%">禮券號碼 

        </td> 

		<td width="37%">

					
<input name="Barcode" type="text" autocomplete="off" value="" size="20" maxlength="15" id="bc" ">


	<input type="Submit" Name="Button" value="確定">


        </td>

     </tr>


</table>

</div>

</form>

<div id="myDiv"></div>

</body>
</HTML>
