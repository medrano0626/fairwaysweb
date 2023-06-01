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

  function checkotherlist1(){
    let form = document.getElementById("form_submit");
    form.submit();
  }

function checkduplicate1(){
    var selected = document.getElementById("otherlist").value;
    var text = document.getElementById("otherlistname").value;
    let form = document.getElementById("form_submit");
    var csrftoken = getCookie('csrftoken');
    
    $.ajax({
        type: "POST",
        url: '/checkduplicate',
        data: { csrfmiddlewaretoken: csrftoken, text: text, selected: selected},
        success: function (data){
                 if(data.otherlistname == "YES"){
                    alert("Name already exist")
                 }else{
                    form.submit();
                 }    
                 }
     });
}

function addtolist(){
if (document.getElementById("otherlist").value == "LOCATION"){
    if (document.getElementById("otherlistname").value == ""){
        alert("Location Name is is empty!");
        return;
      }
}
if (document.getElementById("otherlist").value == "OTHER EARNINGS"){
    if (document.getElementById("otherlistname").value == ""){
        alert("Other Earnings Name is is empty!");
        return;
      }
}
if (document.getElementById("otherlist").value == "VTR EXPENSES"){
    if (document.getElementById("otherlistname").value == ""){
        alert("VTR Expenses Name is is empty!");
        return;
      }
}
document.getElementById("rowid").value="ADD";
checkduplicate1();
}

function removelist(){
    //row = x.rowIndex;
    let form = document.getElementById("form_submit");
    var table = document.getElementById("tablehelper"), rindex, cindex;
    for(var i = 0; i < table.rows.length; i++){
            for(var j = 0; j < table.rows[i].cells.length; j++){
                table.rows[i].cells[j].onclick = function()
                {
                    rindex = this.parentElement.rowIndex;
                     rowid = table.rows[rindex].cells[0].innerHTML
                     let text = "Do you want to delete item?"
                    if (confirm(text) == true) {
                        document.getElementById("rowid").value = rowid
                        form.submit();
                    }
                }

            }
    }
}

function selectdept(){
    var text = document.getElementById("dept").value
    var csrftoken = getCookie('csrftoken');
    $.ajax({
       type: "POST",
       url: 'getperiod',
       data: { csrfmiddlewaretoken: csrftoken, text: text},
       success: function callback(response){
                    $("#period").empty();
                    var period = response.split(',');
                    for(i = 0; i< period.length ; i++){
                      $("#period").append("<option value='" + period[i] + "'>" + period[i] + "</option>");
                    }
                }
    });
  }

  function export_report(){
    let form = document.getElementById("form_submit");
    if(document.getElementById("emptype").value == "EMPTY"){
        alert("Select Employee Type.")
        return;
    }
    if(document.getElementById("payyear1").value == "EMPTY"){
        alert("Select Year.")
        return;
    }
    if(document.getElementById("paymonth1").value == "EMPTY"){
        alert("Select Month.")
        return;
    }
    if(document.getElementById("period").value == "EMPTY"){
        alert("Select Payroll Period.")
        return;
    }
    if(document.getElementById("reporttype").value == "EMPTY"){
        alert("Select Report Type.")
        return;
    }
    if(document.getElementById("location").value == "EMPTY"){
        alert("Select Location.")
        return;
    }
    if(document.getElementById("reporttype").value == "SUMMARY"){
            form.action = "/posted_summary_reports"
            form.submit();
    }else if(document.getElementById("reporttype").value == "PAYSLIP"){
        if(document.getElementById("emptype").value == "OFFICE STAFF" || document.getElementById("emptype").value == "MAINTENANCE"){
            form.action = "/posted_admin_payslip_reports"
            form.submit();
        }else if(document.getElementById("emptype").value == "DRIVER" || document.getElementById("emptype").value == "HELPER"){
            form.action = "/posted_delivery_payslip_reports"
            form.submit();
        }
    }else if(document.getElementById("reporttype").value == "VTR"){
        if(document.getElementById("emptype").value == "DRIVER"){
            form.action = "/posted_driver_vtr_reports"
            form.submit();
        }else if(document.getElementById("emptype").value == "HELPER"){
            form.action = "/posted_helper_vtr_reports"
            form.submit();
        }
    }

  }
  function select_emp_type(){
    var text = document.getElementById("emptype").value
    if(text != "EMPTY"){
        var csrftoken = getCookie('csrftoken');
        $.ajax({
        type: "POST",
        url: 'get_year',
        data: { csrfmiddlewaretoken: csrftoken, text: text},
        success: function callback(response){
                    $("#payyear1").empty();
                    $("#payyear1").append("<option value='EMPTY'>YEAR</option>");
                    $("#paymonth1").empty();
                    $("#paymonth1").append("<option value='EMPTY'>MONTH</option>");
                    $("#period").empty();
                    $("#period").append("<option value='EMPTY'>PERIOD</option>");
                    if(response != ""){
                        var pay_year = response.split(',');
                        for(i = 0; i< pay_year.length ; i++){
                          $("#payyear1").append("<option value='" + pay_year[i] + "'>" + pay_year[i] + "</option>");
                        }  
                    } 
                }
    });
    if(document.getElementById("emptype").value == "OFFICE STAFF" || document.getElementById("emptype").value == "MAINTENANCE"){
        $("#payyear1").empty();
        $("#payyear1").append("<option value='EMPTY'>YEAR</option>");
        $("#paymonth1").empty();
        $("#paymonth1").append("<option value='EMPTY'>MONTH</option>");
        $("#period").empty();
        $("#period").append("<option value='EMPTY'>PERIOD</option>");
        $("#reporttype").empty();
        $("#reporttype").append("<option value='EMPTY'>REPORT TYPE</option>");
        $("#reporttype").append("<option value='SUMMARY'>SUMMARY</option>");
        $("#reporttype").append("<option value='PAYSLIP'>PAYSLIP</option>");
    }else{
        $("#payyear1").empty();
        $("#payyear1").append("<option value='EMPTY'>YEAR</option>");
        $("#paymonth1").empty();
        $("#paymonth1").append("<option value='EMPTY'>MONTH</option>");
        $("#period").empty();
        $("#period").append("<option value='EMPTY'>PERIOD</option>");
        $("#reporttype").empty();
        $("#reporttype").append("<option value='EMPTY'>REPORT TYPE</option>");
        $("#reporttype").append("<option value='SUMMARY'>SUMMARY</option>");
        $("#reporttype").append("<option value='PAYSLIP'>PAYSLIP</option>");
        $("#reporttype").append("<option value='VTR'>VTR</option>");
    }
} 
    
  }
  function select_pay_year(){
    var text = document.getElementById("emptype").value
    var pay_year = document.getElementById("payyear1").value 
    if(pay_year != "EMPTY"){
        var csrftoken = getCookie('csrftoken');
        $.ajax({
        type: "POST",
        url: 'get_month',
        data: { csrfmiddlewaretoken: csrftoken, text: text, pay_year: pay_year},
        success: function callback(response){
                    $("#paymonth1").empty();
                    $("#paymonth1").append("<option value='EMPTY'>MONTH</option>");
                    $("#period").empty();
                    $("#period").append("<option value='EMPTY'>PERIOD</option>");
                    if(response != ""){
                        var pay_month = response.split(',');
                        for(i = 0; i< pay_month.length ; i++){
                          $("#paymonth1").append("<option value='" + pay_month[i] + "'>" + pay_month[i] + "</option>");
                        }  
                    }
                    
                    
                }
    });
    }else{
        $("#paymonth1").empty();
        $("#paymonth1").append("<option value='EMPTY'>MONTH</option>");
        $("#period").empty();
        $("#period").append("<option value='EMPTY'>PERIOD</option>");
    }
    
  }

  function select_pay_month(){
    var text = document.getElementById("emptype").value
    var pay_year = document.getElementById("payyear1").value
    var pay_month = document.getElementById("paymonth1").value
    if(pay_month != "EMPTY"){
        var csrftoken = getCookie('csrftoken');
        $.ajax({
        type: "POST",
        url: 'get_payroll_period',
        data: { csrfmiddlewaretoken: csrftoken, text: text, pay_year: pay_year, pay_month: pay_month},
        success: function (data){
                    $("#period").empty();
                    $("#period").append("<option value='EMPTY'>PERIOD</option>");
                    if(data.period != ""){
                        var pay_period = data.period
                        pay_period = pay_period.split(',');
                        for(i = 0; i< pay_period.length ; i++){
                          $("#period").append("<option value='" + pay_period[i] + "'>" + pay_period[i] + "</option>");
                        }  
                    }
        }
    });
    }else{
        $("#period").empty();
        $("#period").append("<option value='EMPTY'>PERIOD</option>");
    }
    
  }

  function gov_ded_submit(){
    let form = document.getElementById("form_submit");
    if(document.getElementById("payyear1").value == "EMPTY"){
        alert("Select Year.")
        return;
    }
    if(document.getElementById("paymonth1").value == "EMPTY"){
        alert("Select Month.")
        return;
    }
    if(document.getElementById("dedtype").value == "EMPTY"){
        alert("Select Deduction Type.")
        return;
    }
    if(document.getElementById("dedtype").value == "SSS"){
            form.action = "/sss_ded_reports"
            form.submit();
    }else{
            form.action = "/gov_ded_reports"
            form.submit();
        
    }

    

  }

  function ded_submit(){
    let form = document.getElementById("form_submit");
    if(document.getElementById("payyear1").value == "EMPTY"){
        alert("Select Year.")
        return;
    }
    if(document.getElementById("paymonth1").value == "EMPTY"){
        alert("Select Month.")
        return;
    }
    if(document.getElementById("dedtype").value == "EMPTY"){
        alert("Select Deduction Type.")
        return;
    }
   
    form.submit();
    

  }
  function post_payroll(){
        let text = "Posting this payroll period cannot be undone. It is advisable to check all data before posting. If you wish to continue, click OK."
        if (confirm(text) == true) {
        var csrftoken = getCookie('csrftoken');
        var id = "0"
        $.ajax({
        type: "POST",
        url: 'post_payroll',
        data: { csrfmiddlewaretoken: csrftoken, text: id},
        success: function callback(response){
            if (response == "Payroll Period successfully Posted."){
                alert(response)
                window.close();
            }
                    
                }
    });

        }
    

  }

  function deletelist(rowid){ // check vtr if exist, sends command to views.py def testcall(request): to check
    var text = rowid
    let ans = "Do you really want to delete Deduction List?";
    if (confirm(ans) == true){
      var csrftoken = getCookie('csrftoken');
      $.ajax({
         type: "POST",
         url: '/delete_deductions/' + rowid,
         data: { csrfmiddlewaretoken: csrftoken, text: text},
         success: function callback(response){
            location.reload();
                  }
      });}
}

function deletesss(rowid){ // check vtr if exist, sends command to views.py def testcall(request): to check
var text = rowid
let ans = "Do you really want to delete SSS List?";
if (confirm(ans) == true){
  var csrftoken = getCookie('csrftoken');
  $.ajax({
     type: "POST",
     url: '/delete_sss/' + rowid,
     data: { csrfmiddlewaretoken: csrftoken, text: text},
     success: function callback(response){
        location.reload();
              }
  });}
}

function delete_period(rowid){ // check vtr if exist, sends command to views.py def testcall(request): to check
var text = rowid
let ans = "Do you really want to delete Payroll Period List?";
if (confirm(ans) == true){
  var csrftoken = getCookie('csrftoken');
  $.ajax({
     type: "POST",
     url: '/delete_payroll_period/' + rowid,
     data: { csrfmiddlewaretoken: csrftoken, text: text},
     success: function callback(response){
        location.reload();
              }
  });}
}

function edit_period(rowid){
    let form = document.getElementById("form_submit");
    form.action = "/edit_payroll_period/"  + rowid
    form.submit();
    form.action = "/add_payroll_period/0"

}
function close_deductions(){
    let form = document.getElementById("form_submit");
    form.action = "/deductions"
    form.submit();
    form.action = "/add_payroll_period/0"

}
function generate_period(rowid){
    let form = document.getElementById("form_submit");
    form.action = "/generate_payroll/"  + rowid
    console.log("Im here")
    form.submit();
}
function totalsss(){
    if(document.getElementById("ee").value == ""){
        ee = 0
    }else{
        ee = parseFloat(document.getElementById("ee").value)
    }
    if(document.getElementById("er").value == ""){
        er = 0
    }else{
        er = parseFloat(document.getElementById("er").value)
    }
    if(document.getElementById("ec").value == ""){
        ec = 0
    }else{
        ec = parseFloat(document.getElementById("ec").value)
    }
    if(document.getElementById("wispee").value == ""){
        wispee = 0
    }else{
        wispee = parseFloat(document.getElementById("wispee").value)
    }
    if(document.getElementById("wisper").value == ""){
        wisper = 0
    }else{
        wisper = parseFloat(document.getElementById("wisper").value)
    }
    
    total_ee = ee + wispee
    total_er = er + ec + wisper
    
    document.getElementById("totee").value = total_ee
    document.getElementById("toter").value = total_er
    }
    
function check_selected_reports(id){
    let form = document.getElementById("form_submit");
    form.action = "/other_reports/"  + id
    form.submit();
}

function export_reports1(){
    let form = document.getElementById("form_submit");
    if(document.getElementById("emptype").value == "EMPTY"){
        alert("Select Employee Type.")
        return;
    }
    if(document.getElementById("datefrom").value == ""){
        alert("Input start date.")
        return;
    }
    if(document.getElementById("dateto").value == ""){
        alert("Input end date.")
        return;
    }
    if(document.getElementById("location").value == "EMPTY"){
        alert("Select Locataion.")
        return;
    }
    // form.submit();
}

function submit_upload(id){
    let form = document.getElementById("form_submit");
    form.action = "/submit_upload/"  + id
    form.submit();
}
function submit_fix(id){
    let form = document.getElementById("form_submit");
    form.action = "/submit_fix/"  + id
    form.submit();
}
