import {
    setLoggedinEmployeeId
 } from './modules/northwindIdentityService.js';

 const button1Element = document.getElementById('loginButton1');
 button1Element.addEventListener('click', ev => {
    setLoggedinEmployeeId(1);
    window.location.href = "/";
 });

 const button2Element = document.getElementById('loginButton2');
 button2Element.addEventListener('click', ev => {
    setLoggedinEmployeeId(2);
    window.location.href = "/";
 });