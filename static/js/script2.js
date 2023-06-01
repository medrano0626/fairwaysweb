var mylistmonth = document.getElementById("paymonth1").options;
var mylistyear = document.getElementById("payyear1").options;
var paymonth1 = document.getElementById('paymonth1');
var payyear1 = document.getElementById('payyear1');
const data = document.currentScript.dataset;
const paymonthval = data.paymonth;
const payyearval = data.payyear;

yearnow = new Date().getFullYear();
yearlast = yearnow - 1
yearnext = yearnow + 1

if (document.getElementById("periodto").defaultValue == ""){
    document.getElementById("periodto").defaultValue = new Date().toISOString().substring(0, 10);
}

if (document.getElementById("periodfrom").defaultValue == "") {
    document.getElementById("periodfrom").defaultValue = new Date().toISOString().substring(0, 10);                 
}

for (var i = 1; i <= 12; i++) {
    paymonth1.add(new Option(i));
}

for (var i = yearlast; i <= yearnext; i++) {
    payyear1.add(new Option(i));
}
   
for (var i = 0; i < mylistmonth.length; i++) {
    if (mylistmonth[i].text == paymonthval) {
        mylistmonth[i].selected = true;
        break;
    }
    
}

for (var i = 0; i < mylistyear.length; i++) {
    if (mylistyear[i].text == payyearval) {
        mylistyear[i].selected = true;
        break;
    }
    
}


