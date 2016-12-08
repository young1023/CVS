<%

Response.write Session("SecLevel")

'Redirect to login page if timeout or no security session

If Session("SecLevel") < 3 Then
   
   Response.Redirect "Default.asp"

End if

%>

<table width="100%" border="0" cellspacing="2" cellpadding="2">
  <tr>
    <td height="72" width="10%"><IMG style="LEFT: 0px; POSITION: absolute; TOP: 0px"  alt ="Shell logo" src="IMAGES/Shell.gif" border=0 ></IMG></td>
   
<td align="right">

<table width="100%" border="0" cellspacing="2" cellpadding="4">
  <tr>

<% If Session("SecLevel") = 3 or Session("SecLevel") > 5 then %> 
   
<td align="right">
<a href="pco1.asp" title="Create new procoupon" style="TEXT-DECORATION: none">
New Procoupon</a>
</td>

<td align="right">
<a href="pco_d1.asp" title="Procoupon" style="TEXT-DECORATION: none">
Procoupon</a>
</td>

<% End If %>

<% If Session("SecLevel") > 5 then %> 

<td align="right">
<a href="master1.asp" title="Master List" style="TEXT-DECORATION: none">
Master List</a>
</td>
<td align="right">
<a href="rtn_co1.asp" title="Return coupon" style="TEXT-DECORATION: none">
Return Coupon</a>
</td>

<% End If %>

<% If Session("SecLevel") >= 5 then %> 


<td align="right">
<a href="rd_co2.asp" title="Redemption coupon" style="TEXT-DECORATION: none">
 Redemption Coupon</a>
</td>
<td align="right">
<a href="rd_r_co1.asp" title="Remdemtion raw coupon" style="TEXT-DECORATION: none">
Redemption Raw Coupon</a>
</td>
<td align="right">
<a href="wk_rpt1.asp" title="Weekly report" style="TEXT-DECORATION: none">
Weekly Report</a>
</td>

<% End If %>

<% If Session("SecLevel") > 5 then %> 

<td align="right">
<a href="station1.asp" title="Station" style="TEXT-DECORATION: none">
Station</a>
</td>
<td align="right">
<a href="user1.asp" title="User" style="TEXT-DECORATION: none">
User</a>
</td>
<td align="right">
<a href="coupon1.asp" title="Station" style="TEXT-DECORATION: none">
Coupon</a>
</td>

<% End If %>

<td align="right">
<a href="default.asp" title="Station" style="TEXT-DECORATION: none">
Logout</a>
</td>
  </tr>
</table>


</td>
  </tr>
</table>
