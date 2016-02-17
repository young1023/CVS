<SCRIPT language=JavaScript>
<!--

function doLogin(){

   document.fm1.action="Login.asp";
     
   if (document.fm1.UserID.value == "")
{
   alert("請輸入用戶編號.");
   document.fm1.UserID.focus();
   return false;
   }

 if (document.fm1.Password.value == "")
  
{
   alert("請輸入密碼.");
   document.fm1.Password.focus();
   return false;
   }	
   document.fm1.submit();
}



// Check digit
function mod10CheckDigit() {
//Get barCode field value, Strip all characters
//except numbers Repopulate field w/ new value
bc = document.barCodeForm.Barcode.value
checkdigit = bc.substr(3,1);
bc = bc.replace(/[^0-9]+/g,'');
//bc = bc.replace(/^\s+|\s+$/gm,'');
document.barCodeForm.Barcode.value = bc;
total = 0;
checkdigit2 = null;

// if coupon has 16 digits
if (bc.length==16){

// get the check digit at 16th digit
checkdigit2 = bc.substr(15,1);
// change the coupon length back to 15th digits for the calculation
bc = bc.substr(0,15)

}

//Get Even Numbers
for (i=bc.length-2; i>=0; i=i-2) {
if (i != 3 ){
tempeven = parseInt(bc.substr(i,1)) ;
if (tempeven > 9) {
tempeven = tempeven - 9;
}
total = total + tempeven ;
}
}
//Get Odd Numbers
for (i=bc.length-1; i>=0; i=i-2) {
temp = parseInt(bc.substr(i,1)) * 2 ;
if (temp > 9) {
temp = temp - 9;
}
total = total + temp;
}

//Determine the checksum. if modDigit = 10, modDigit = 0
modDigit = (10 - ((total + 17) % 10)) % 10;


if (checkdigit != modDigit){
alert("驗證失敗 - 禮券無效!");
document.barCodeForm.Barcode.value = "";
document.barCodeForm.Barcode.focus();
return false;

}


if (checkdigit2 !== null){
//Determine the 2nd checksum. if modDigit = 10, modDigit = 0
modDigit2 = (10 - ((total + 23) % 10)) % 10;
// if the 2nd check digit does not equal
if (checkdigit2 != modDigit2){
alert("驗證失敗 - 禮券無效!");
document.barCodeForm.Barcode.value = "";
document.barCodeForm.Barcode.focus();
return false;
}
}


//Populate the modDigit field with the value
document.barCodeForm.action="execute.asp";
document.barCodeForm.submit();
}



//-->
</SCRIPT>