var mylistdedtype = document.getElementById("dedtype").options;
var mylistperiod = document.getElementById("payrollperiod").options;
const data = document.currentScript.dataset;
const dedtypevalue = data.dedtypevalue;
const periodvalue = data.periodvalue;

for (var i = 0; i < mylistdedtype.length; i++) {
    if (mylistdedtype[i].text == dedtypevalue) {
        mylistdedtype[i].selected = true;
        break;
    }
    
}

for (var i = 0; i < mylistperiod.length; i++) {
    if (mylistperiod[i].text == periodvalue) {
        mylistperiod[i].selected = true;
        break;
    }
    
}
