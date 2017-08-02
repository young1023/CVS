<HTML>
<HEAD>
<TITLE>Barcode Image</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" type="text/css" href="include/publish.css" />
<SCRIPT language=JavaScript>
<!--
function initializeMainDiv() {
        document.getElementById("mainDiv").innerHTML = "Barcode: " +
            window.opener.document.getElementById("Barcode3").value
    }


function myFunction() {
    window.print();
}
//-->
</SCRIPT>
</HEAD>
<body>

<div align = "center">

<button class="noprint" onclick="myFunction()">Print this page</button>

<br/>

<br/>

<%

i = 0

y = 1

barcode = request("barcode")

'response.write barcode

' calculate number of barcode received
x = len(barcode)/16

Do while x > i


' Generate image from url

%>

<p>

<img src='https://barcode.tec-it.com/barcode.ashx?translate-esc=off&data=<% = Mid(barcode, y  ,16) %>&code=Code128&unit=Fit&dpi=96&imagetype=Gif&rotation=0&color=000000&bgcolor=FFFFFF&qunit=Mm&quiet=0' />

</p>

<br/>

<br/>


<%


y = y + 16



i = i + 1

Loop



%>


</div>

</BODY>
</HTML>
