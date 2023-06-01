var myDLInput = document.getElementById('driver');
var mySelect = document.getElementById("drivername");
const dinput = document.getElementById('driver');
let eventSource = null;
let value = '';
var myDLInput1 = document.getElementById('helper');
var mySelect1 = document.getElementById("helpername");
const dinput1 = document.getElementById('helper');
var mylistperiod = document.getElementById("payrollperiod").options;
const data = document.currentScript.dataset;
const oldvtrno = data.oldvtrno;
const oldperiod = data.oldperiod;
let form = document.getElementById("form_submit");



dinput1.addEventListener('keydown', (e) => {
  eventSource = e.key ? 'input' : 'list';
});


dinput1.addEventListener('input', (e) => {
  value = e.target.value;
  if (eventSource === 'list') {
        document.getElementById("helpername").value = value
        let selectedOptionID = mySelect1.options[mySelect1.selectedIndex].dataset.id;
        document.getElementById("helperempno").value = selectedOptionID
        empno = selectedOptionID
        getemprate(empno,0)
  }
});


dinput.addEventListener('keydown', (e) => {
  eventSource = e.key ? 'input' : 'list';
});


dinput.addEventListener('input', (e) => {
  value = e.target.value;
  if (eventSource === 'list') {
        document.getElementById("drivername").value = value
        let selectedOptionID = mySelect.options[mySelect.selectedIndex].dataset.id;
        document.getElementById("driverempno").value = selectedOptionID
        empno = selectedOptionID
        getemprate(empno,1)

  }
});

function SomeDeleteRowFunction() { // delete rows in table
  var td = event.target.parentNode; 
  var tr = td.parentNode; 
  tr.parentNode.removeChild(tr);
  helper = "";
  var allhelper = "('m','m','m','m','m','m')"
    var table = document.getElementById('tablehelper');
    for (var r = 1, n = table.rows.length; r < n; r++) {
      helper = "('" + table.rows[r].cells[0].innerHTML + "','" + table.rows[r].cells[1].innerHTML + "','" + table.rows[r].cells[2].innerHTML + "','" + table.rows[r].cells[3].innerHTML + "','" + table.rows[r].cells[4].innerHTML + "','" + table.rows[r].cells[5].innerHTML + "')"
        allhelper = allhelper + "," + helper
    }
    document.getElementById("helpers").value = allhelper
    
    var table = document.getElementById("tablexpenses");
    var allexpenses = "('m','m')";
      expenses = "";
      for (var r = 1, n = table.rows.length; r < n; r++) {
        expenses = "('" + table.rows[r].cells[0].innerHTML + "','" + table.rows[r].cells[1].innerHTML + "')"
        allexpenses = allexpenses + "," + expenses
      }
      document.getElementById("expenseslist").value = allexpenses
      
}


function addhelper() { // add helper button
  if(document.getElementById("helperempno").value != "" && document.getElementById("helpertripamount").value != ""){
    var table = document.getElementById("tablehelper");
    var helper = "";
    var totalrow = table.rows.length
    var row = table.insertRow(totalrow);
    var cell1 = row.insertCell(0);
    var cell2 = row.insertCell(1);
    var cell3 = row.insertCell(2);
    var cell4 = row.insertCell(3);
    var cell5 = row.insertCell(4);
    var cell6 = row.insertCell(5);
    var cell7 = row.insertCell(6);
    cell1.innerHTML = document.getElementById("helperempno").value;
    cell2.innerHTML = document.getElementById("helpername").value;
    cell3.innerHTML = document.getElementById("helpertripamount").value;
    if(document.getElementById("helperallowance").value == ""){
      cell4.innerHTML = "0"
    }
    if(document.getElementById("helperallowance").value != ""){
      cell4.innerHTML = document.getElementById("helperallowance").value;
    }
    if(document.getElementById("helperexcesstrip").value == ""){
      cell5.innerHTML = "0"
    }
    if(document.getElementById("helperexcesstrip").value != ""){
      cell5.innerHTML = document.getElementById("helperexcesstrip").value;
    }
    if(document.getElementById("helpertruckvale").value == ""){
      cell6.innerHTML = "0"
    }
    if(document.getElementById("helpertruckvale").value != ""){
      cell6.innerHTML = document.getElementById("helpertruckvale").value;
    } 
    cell7.innerHTML = "<input class = 'removehelper' type='button' value='Remove' onclick='SomeDeleteRowFunction()'></input>"
    document.getElementById("helper").value = ""
    document.getElementById("helperempno").value = ""
    document.getElementById("helpertripamount").value = ""
    document.getElementById("helperexcesstrip").value = ""
    document.getElementById("helpertruckvale").value = ""
    document.getElementById("helperallowance").value = ""
    var allhelper = "('m','m','m','m','m','m')"
    var table = document.getElementById('tablehelper');
    for (var r = 1, n = table.rows.length; r < n; r++) {
      helper = "('" + table.rows[r].cells[0].innerHTML + "','" + table.rows[r].cells[1].innerHTML + "','" + table.rows[r].cells[2].innerHTML + "','" + table.rows[r].cells[3].innerHTML + "','" + table.rows[r].cells[4].innerHTML + "','" + table.rows[r].cells[5].innerHTML + "')"
        
        allhelper = allhelper + "," + helper
    }
    document.getElementById("helpers").value = allhelper
    

  }

}


  function addexpenses() { // add expenses button
    if (document.getElementById("expensesamount").value != "") {
      var table = document.getElementById("tablexpenses");
      var totalrow = table.rows.length
      var row = table.insertRow(totalrow);
      var cell1 = row.insertCell(0);
      var cell2 = row.insertCell(1);
      var cell3 = row.insertCell(2);
      var allexpenses = "('m','m')";
      expenses = "";
      cell1.innerHTML = document.getElementById("expenses").value;
      cell2.innerHTML = document.getElementById("expensesamount").value;
      cell3.innerHTML = "<input class = 'removehelper' type='button' value='Remove' onclick='SomeDeleteRowFunction()'></input>"
      document.getElementById("expensesamount").value = ""
      for (var r = 1, n = table.rows.length; r < n; r++) {
        expenses = "('" + table.rows[r].cells[0].innerHTML + "','" + table.rows[r].cells[1].innerHTML + "')"
        allexpenses = allexpenses + "," + expenses
      }
      document.getElementById("expenseslist").value = allexpenses
      
  
    }
    
}
function submitForm() {
  let vtrdate = document.getElementById("vtrdate").value;
  var period = document.getElementById("payrollperiod").value;
  let periodfrom = period.slice(0,10);
  let periodto = period.slice(11,21);
  vtrdate = new Date(vtrdate);
  periodfrom = new Date(periodfrom);
  periodto = new Date(periodto);
  if (document.getElementById("payrollperiod").value == ""){
    alert("Payroll period is empty!");
    return;
  }
  if (document.getElementById("vtrno").value == ""){
    alert("VTRNO is empty!");
    return;
  }
  if (document.getElementById("vtrdate").value == ""){
    alert("VTRDATE is empty!");
    return;
  }
  if (vtrdate < periodfrom || vtrdate > periodto){
    alert("VTR Date is not under the payroll period you selected!");
    return;
  }
  if (document.getElementById("plateno").value == ""){
    alert("PLATE NO is empty!");
    return;
  }
  if (document.getElementById("locationfrom").value == ""){
    alert("LOCATION FROM is empty!");
    return;
  }
  if (document.getElementById("locationto").value == ""){
    alert("LOCATION TO is empty!");
    return;
  }
  if (document.getElementById("driverempno").value == ""){
    alert("DRIVER is empty!");
    return;
  }
  let text = "This will update the VTR record. Do you want to continue?";
  if (confirm(text) == true) {
    if (oldperiod != document.getElementById("payrollperiod").value){
        let text = "Payroll Period changed. This VTR will be transferred to the period you selected and may cause conflict (if the payroll period is posted) to the last period it was saved.";
        if (confirm(text) == true){
            if (oldvtrno != document.getElementById("vtrno").value){
              checkvtr(); 
              }
        if(oldvtrno == document.getElementById("vtrno").value){
            console.log("submit form")
            form.submit();
              }
        }

    } 
    if (oldperiod == document.getElementById("payrollperiod").value){
        if (oldvtrno != document.getElementById("vtrno").value){
          checkvtr(); 
        }
        if(oldvtrno == document.getElementById("vtrno").value){
          console.log("submit form")
          form.submit();
        }
  }
  
}
}

for (var i = 0; i < mylistperiod.length; i++) {
  if (mylistperiod[i].text == oldperiod) {
      mylistperiod[i].selected = true;
      break;
  }
  
}


function getCookie(name) {
  var cookieValue = null;
  if (document.cookie && document.cookie !== '') {
      var cookies = document.cookie.split(';');
      for (var i = 0; i < cookies.length; i++) {
          var cookie = cookies[i].trim();
          // Does this cookie string begin with the name we want?
          if (cookie.substring(0, name.length + 1) === (name + '=')) {
              cookieValue = decodeURIComponent(cookie.substring(name.length + 1));
              break;
          }
      }
  }
  return cookieValue;
}

function checkvtr(){ // check vtr if exist, sends command to views.py def testcall(request): to check
  var text = document.getElementById('vtrno').value;
        // let form = document.getElementById("form_submit");
        var csrftoken = getCookie('csrftoken');

        $.ajax({
           type: "POST",
           url: '/testvtr',
           data: { csrfmiddlewaretoken: csrftoken, text: text},
           success: function callback(response){
                       console.log(response);
                       if(response == "YES"){
                            alert("VTRNO already exist!")
                       }
                       if(response == "NO"){
                        form.submit();
                      }
                        
                        
                    }
        });
}
function deletevtr(){
  if(document.getElementById('vtrno').value != ""){
    let text = "Do you really want to delete VTR?";
        if (confirm(text) == true){
          document.getElementById("form_submit").action = "/delete_vtr/" +  document.getElementById('vtrno').value
          form.submit();
        }
    
  }

}
function getemprate(empno,a){ 
  var text = empno
        var csrftoken = getCookie('csrftoken');

        $.ajax({
           type: "POST",
           url: '/getrate',
           data: { csrfmiddlewaretoken: csrftoken, text: text},
           success: function callback(response){
            if(a == 0){
              document.getElementById("helpertripamount").value = response  
            }
            if(a == 1){
              document.getElementById("drivertripamount").value = response  
            }
                 
                        
                    }
        });
        
}