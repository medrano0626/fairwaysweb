const data = document.currentScript.dataset;
const image_data = data.image_data;     
if(image_data != "None"){
    var base64img = image_data.slice(2, image_data.length-1);
    function Base64ToImage(base64img, callback) {
        var img = new Image();
        img.onload = function() {
            callback(img);
        };
        img.src = base64img;
    }
    Base64ToImage(base64img, function(img) {
        document.getElementById('results1').appendChild(img);
        document.getElementById("results1").style.display = "block"
        document.getElementById("results").style.display = "none"
    });
}else{
    document.getElementById("results1").style.display = "none"
    document.getElementById("results").style.display = "block"
   
}


Webcam.set({
    width: 350,
    height: 287,
    image_format: 'jpeg',
    jpeg_quality: 90
    });	 
    Webcam.attach( '#my_camera' );
    
    function take_snapshot() {
    // play sound effect
    //shutter.play();
    // take snapshot and get image data
    Webcam.snap( function(data_uri) {
    // display results in page
    document.getElementById('results').innerHTML = 
    '<img style="width: 350px; height: 275px;" class="after_capture_frame" src="'+data_uri+'"/>';
    document.getElementById('results1').innerHTML = 
    '<img style="width: 350px; height: 275px;" class="after_capture_frame" src="'+data_uri+'"/>';
    $("#captured_image_data").val(data_uri);
    });	 
    }

    function saveSnap(){
    var base64data = $("#captured_image_data").val();
    if (base64data == ""){
        alert("No image captured.")
    }else{
        let form = document.getElementById("form_submit");
        form.submit();
    }
    }        
