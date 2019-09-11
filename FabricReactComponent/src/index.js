"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
// These are the references to the react library
var React = require("react");
var ReactDOM = require("react-dom");
var EmployeeCardLayout_1 = require("./EmployeeCardLayout");
// Get the "main" element
var target = document.querySelector("#EmployeeCardMain");
if (target) {
    ReactDOM.render(React.createElement(EmployeeCardLayout_1.EmployeeCardLayout, null), target);
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
//# sourceMappingURL=index.js.map