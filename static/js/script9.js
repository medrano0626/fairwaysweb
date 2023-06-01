let form = document.getElementById("form_submit");
var rowid;



document.getElementById("datemaintained").defaultValue = new Date().toISOString().substring(0, 10);
document.getElementById("datelastupdate").defaultValue = new Date().toISOString().substring(0, 10);




function adddeductions(){

    if (document.getElementById("amount").value == ""){
      alert("Amount is empty!");
      return;
    }
    if (document.getElementById("period").value == ""){
      alert("Payroll Period is empty!");
      return;
    }
    form.submit();
}

function updatedeductions(){

    if (document.getElementById("amount").value == ""){
        alert("Amount is empty!");
        return;
      }
      if (document.getElementById("period").value == ""){
        alert("Payroll Period is empty!");
        return;
      }
    form.submit();
}
function addearnings(){

    if (document.getElementById("amount").value == ""){
      alert("Amount is empty!");
      return;
    }
    if (document.getElementById("period").value == ""){
      alert("Payroll Period is empty!");
      return;
    }
    form.submit();
}

function updateearnings(){

    if (document.getElementById("amount").value == ""){
        alert("Amount is empty!");
        return;
      }
      if (document.getElementById("period").value == ""){
        alert("Payroll Period is empty!");
        return;
      }
    form.submit();
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
function checkdeductions(){ // check vtr if exist, sends command to views.py def testcall(request): to check
                var text = rowid;
              var csrftoken = getCookie('csrftoken');
              $.ajax({
                 type: "POST",
                 url: '/checkdeductions1',
                 data: { csrfmiddlewaretoken: csrftoken, text: text},
                 success: function (data){//amount, noofperiod, dedamount, balance){
                    document.getElementById("period").value = data.period
                    document.getElementById("dedtype").value = data.dedtype
                    document.getElementById("amount").value = data.dedamount
                    document.getElementById("datemaintained").value = data.datemaintained
                    document.getElementById("datelastupdate").value = data.datelastupdate
                    document.getElementById("user").value = data.user           
                    document.getElementById("rowid").value = rowid
                    document.getElementById("addded").style.display = "none"
                    document.getElementById("updateded").style.display = "block"
                          }
              });
      }
function editdeductions1(){
    //row = x.rowIndex;
    var table = document.getElementById("tablehelper"), rindex, cindex;
    for(var i = 0; i < table.rows.length; i++){
            for(var j = 0; j < table.rows[i].cells.length; j++){
                table.rows[i].cells[j].onclick = function()
                {
                    rindex = this.parentElement.rowIndex;
                     rowid = table.rows[rindex].cells[0].innerHTML
                    checkdeductions();
                }

            }
    }
    

}

function editearnings1(){
    //row = x.rowIndex;
    var table = document.getElementById("tablehelper"), rindex, cindex;
    for(var i = 0; i < table.rows.length; i++){
            for(var j = 0; j < table.rows[i].cells.length; j++){
                table.rows[i].cells[j].onclick = function()
                {
                    rindex = this.parentElement.rowIndex;
                     rowid = table.rows[rindex].cells[0].innerHTML
                    checkearnings();
                }

            }
    }
    

}

function deletedeductions1(){
    //row = x.rowIndex;

    var table = document.getElementById("tablehelper"), rindex, cindex;
    for(var i = 0; i < table.rows.length; i++){
            for(var j = 0; j < table.rows[i].cells.length; j++){
                table.rows[i].cells[j].onclick = function()
                {
                    rindex = this.parentElement.rowIndex;
                     rowid = table.rows[rindex].cells[0].innerHTML
                     let text = "Do you want to delete deduction?";
                    if (confirm(text) == true) {
                        document.getElementById("rowid").value = rowid
                        document.getElementById("rowid1").value = rowid
                        form.submit();
                    }
                }

            }
    }

    
    

}
function deleteearnings1(){
    //row = x.rowIndex;

    var table = document.getElementById("tablehelper"), rindex, cindex;
    for(var i = 0; i < table.rows.length; i++){
            for(var j = 0; j < table.rows[i].cells.length; j++){
                table.rows[i].cells[j].onclick = function()
                {
                    rindex = this.parentElement.rowIndex;
                     rowid = table.rows[rindex].cells[0].innerHTML
                     let text = "Do you want to delete deduction?";
                    if (confirm(text) == true) {
                        document.getElementById("rowid").value = rowid
                        document.getElementById("rowid1").value = rowid
                        form.submit();
                    }
                }

            }
    }

    
    

}
function checkearnings(){ // check vtr if exist, sends command to views.py def testcall(request): to check
    var text = rowid;
  var csrftoken = getCookie('csrftoken');
  $.ajax({
     type: "POST",
     url: '/checkearnings1',
     data: { csrfmiddlewaretoken: csrftoken, text: text},
     success: function (data){//amount, noofperiod, dedamount, balance){
        document.getElementById("period").value = data.period
        document.getElementById("dedtype").value = data.dedtype
        document.getElementById("amount").value = data.dedamount
        document.getElementById("datemaintained").value = data.datemaintained
        document.getElementById("datelastupdate").value = data.datelastupdate
        document.getElementById("user").value = data.user           
        document.getElementById("rowid").value = rowid
        document.getElementById("addded").style.display = "none"
        document.getElementById("updateded").style.display = "block"
              }
  });
}
function backtoempded(){
    document.getElementById("form_submit").action = "/emp_deductions"
    form.submit();
}

  
