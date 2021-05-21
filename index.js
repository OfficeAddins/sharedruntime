window.onload = function Start() { 
    console.log('hello world'); 
 
    var app_1 = document.getElementById("app"); 
 
    app_1.innerHTML = app_1.innerHTML + 
       '<br><input type="radio" id="radiobutton1" name="mygroup"><label for="radiobutton1">Monday</label>' +
       '<br><input type="radio" id="radiobutton2" name="mygroup"><label for="radiobutton2">Tuesday</label>' +
       '<br><input type="radio" id="radiobutton3" name="mygroup"><label for="radiobutton3">Wednesday</label>';
 } 
 
 Office.initialize = function (reason) { 
 } 

 window.btnMonday = btnMonday;
 function btnMonday(event) {
    var radio1 = document.getElementById("radiobutton1").checked = true;
    event.completed();
 }
 
 window.btnTuesday = btnTuesday;
 function btnTuesday(event) {
    var radio1 = document.getElementById("radiobutton2").checked = true;
    event.completed();
 }

 window.btnWednesday = btnWednesday;
 function btnWednesday(event) {
    var radio1 = document.getElementById("radiobutton3").checked = true;
    event.completed();
 }