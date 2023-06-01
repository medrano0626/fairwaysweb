
        var form_1 = document.querySelector(".form.first");
        var form_2 = document.querySelector(".form.second");
        var form_3 = document.querySelector(".form.third");
        var form_1_next_btn = document.querySelector(".nextBtn1");
        var form_2_back_btn = document.querySelector(".backBtn1");
        var form_2_next_btn = document.querySelector(".nextBtn2");
        var form_3_back_btn = document.querySelector(".backBtn2");
        var btn_done = document.querySelector(".submit");
        var modal_wrapper = document.querySelector(".modal_wrapper");
        var shadow = document.querySelector(".shadow");
        var mylistlocation = document.getElementById("location").options;
        const data = document.currentScript.dataset;
        const locationval = data.location;


        form_1_next_btn.addEventListener("click", function(){
            
            form_1.classList.remove("active");
            form_2.classList.add("active");
            form_3.classList.remove("active");
            // form_4.classList.remove("active");
            document.getElementById("sss").focus();
            
        });

        form_2_back_btn.addEventListener("click", function(){
            form_1.classList.add("active");
            form_2.classList.remove("active");
            form_3.classList.remove("active");
            // form_4.classList.remove("active");
            document.getElementById("lname").focus();
        });

        form_2_next_btn.addEventListener("click", function(){
            form_1.classList.remove("active");
            form_2.classList.remove("active");
            form_3.classList.add("active");
            // form_4.classList.remove("active");
            document.getElementById("triprate").focus();
    
        });

        form_3_back_btn.addEventListener("click", function(){
            form_1.classList.remove("active");
            form_2.classList.add("active");
            form_3.classList.remove("active");
            // form_4.classList.remove("active");
            document.getElementById("sss").focus();
            
        });

        btn_done.addEventListener("click", function(){
            
        })
        function getdate(){
             if (document.getElementById("birthdate").defaultValue == ""){
                document.getElementById("birthdate").defaultValue = new Date().toISOString().substring(0, 10);
             }
             if (document.getElementById("datehired").defaultValue == "") {
                document.getElementById("datehired").defaultValue = new Date().toISOString().substring(0, 10);                 
             }
             if (document.getElementById("datepermanent").defaultValue  == ""){
                document.getElementById("datepermanent").defaultValue = new Date().toISOString().substring(0, 10);
             }
             
            } 


for (var i = 0; i < mylistlocation.length; i++) {
        if (mylistlocation[i].text == locationval) {
            mylistlocation[i].selected = true;
                break;
        }
                
}