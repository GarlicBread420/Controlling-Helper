!function(){"use strict";var t={80444:function(t,e,n){t.exports=n.p+"5d97eaa7d68c3a557b63.ts"},13014:function(t,e,n){t.exports=n.p+"0a0e74765913d8f5fc8a.css"}},e={};function n(o){var r=e[o];if(void 0!==r)return r.exports;var c=e[o]={exports:{}};return t[o](c,c.exports,n),c.exports}n.m=t,n.g=function(){if("object"==typeof globalThis)return globalThis;try{return this||new Function("return this")()}catch(t){if("object"==typeof window)return window}}(),n.o=function(t,e){return Object.prototype.hasOwnProperty.call(t,e)},function(){var t;n.g.importScripts&&(t=n.g.location+"");var e=n.g.document;if(!t&&e&&(e.currentScript&&"SCRIPT"===e.currentScript.tagName.toUpperCase()&&(t=e.currentScript.src),!t)){var o=e.getElementsByTagName("script");if(o.length)for(var r=o.length-1;r>-1&&(!t||!/^http(s?):/.test(t));)t=o[r--].src}if(!t)throw new Error("Automatic publicPath is not supported in this browser");t=t.replace(/#.*$/,"").replace(/\?.*$/,"").replace(/\/[^\/]+$/,"/"),n.p=t}(),n.b=document.baseURI||self.location.href,function(){function t(){}t.getMonthArray=function(){return["Januar","Februar","März","April","Mai","Juni","Juli","August","September","Oktober","November","Dezember"]},t.populateDropdown=function(t,e){const n=document.getElementById(e);t.forEach((t=>{const e=document.createElement("option");e.value=t,e.text=t,n.appendChild(e)}))},t.goBack=function(){window.location.href="../taskpane.html"},t.selectAll=function(t,e){var n=document.getElementById(t);if(n instanceof HTMLInputElement){let t=document.getElementsByName(e);for(let e=0;e<t.length;e++){const o=t[e];o instanceof HTMLInputElement&&(o.checked=n.checked)}}},Office.onReady((e=>{e.host===Office.HostType.Excel&&(document.getElementById("backButton").onclick=t.goBack)}))}(),new URL(n(13014),n.b),new URL(n(80444),n.b)}();
//# sourceMappingURL=autoControlling_feature.js.map