<% 
option explicit 
response.expires = 0
response.buffer = true
%>
<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN"">
<HTML>
<head>
<%
if Request.QueryString("mode") = "year" then
%>
	<TITLE> Year Calendar </TITLE>
<%
else
%>
	<TITLE> Month Calendar </TITLE>
<%
end if
%>
</head>
<body>
<center>
	<table>
		<tr>
			<td>
<%
dim cal
set cal = new calendar
if Request.QueryString("mode") = "year" then
		cal.yr = Request.QueryString("year")
		Response.Write( Cal.YearCal )
else
	cal.FullYearLink = "test.asp?mode=year"
	Response.Write( Cal.MonthCal )
end if
set cal = nothing
%>		</td>
		</tr>
	</table>
</center>
</body>
</html>
<!-- #INCLUDE FILE="calendar.asp" -->

