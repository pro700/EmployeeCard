// These are the references to the react library
import * as React from "react";
import * as ReactDOM from "react-dom";

import * as $ from 'jquery';
import { EmployeeCardLayout } from "./EmployeeCardLayout";


// Get the "main" element
let target = document.querySelector("#EmployeeCardMain");
if (target) {

    ReactDOM.render(<EmployeeCardLayout/>, target);

}

window.addEventListener("resize", onResize);

function onResize() {
    var w = document.documentElement.clientWidth;
    var h = document.documentElement.clientHeight;

    var regex = new RegExp(/[Ss]ender[Ii]d=([\daAbBcCdDeEfF]+)/);
    var results = regex.exec(window.location.search);

    if (null != results && null != results[1]) {
        window.parent.postMessage('<message senderId=' + results[1] + '>resize(' + w + ',' + h + ')</message>', '*');
    }

}