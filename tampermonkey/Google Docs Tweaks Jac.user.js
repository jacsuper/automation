// ==UserScript==
// @name Google Docs Tweaks Jac
// @namespace some_google_doc_shortcuts_to_change_font_highlight
// @version 1.0.2
// @author Willi, Jac
// @description some google doc shortcuts to change the font and highlights (help needed for font size!). It also hide the page break dotted lines.

// @match    https://docs.google.com/document/*
// @require    http://ajax.googleapis.com/ajax/libs/jquery/1/jquery.min.js

// @grant    GM_addStyle

// ==/UserScript==
// hide the page break doted line:
// https://stackoverflow.com/a/52870899/3154274
GM_addStyle ( `

.kix-page-compact::before {
        border-top: none;
    }

// color in red current title in outline
.outline-refresh .location-indicator-highlight.navigation-item,  .outline-refresh .location-indicator-highlight.navigation-item.goog-button-hover {
color: rgb(255, 0, 0);
}


` );


 // listen for key shorcuts on the text part of google gocs
//	● Tuto - sources:
// 		○ for iframe https://stackoverflow.com/a/46217408/3154274
// 		○ for switch https://stackoverflow.com/q/13362028/3154274
// 		○ combinaison of key  https://stackoverflow.com/a/37559790/3154274
// 		○ dispatchEvent https://stackoverflow.com/a/33887557/3154274

// 		○ for dispatch :
//		    https://jsfiddle.net/6vyL98mz/33/
//    		https://jsfiddle.net/ox2La621/1/

var editingIFrame = $('iframe.docs-texteventtarget-iframe')[0];
if (editingIFrame) {
    editingIFrame.contentDocument.addEventListener("keydown", dispatchkeyboard, false);
}
else {
    setTimeout(function(){
        var editingIFrame = $('iframe.docs-texteventtarget-iframe')[0];
        if (editingIFrame) {
            editingIFrame.contentDocument.addEventListener("keydown", dispatchkeyboard, false);
        }
    }, 3);
}



// match the key with the action
function dispatchkeyboard(key) {
    //--------------- background color
    var buttonbg;
    // hl yellow
    if ((key.altKey && key.code === "KeyY") || (key.shiftKey && key.ctrlKey && key.code === "KeyY")) {
        buttonbg = document.getElementById("bgColorButton");

        callMouseEvent(buttonbg);
        setTimeout(function(){
            var color_choice = document.getElementById("docs-material-colorpalette-cell-113");//buttonbg.querySelector('[title="yellow"]');
            console.log("clickbutton wait 2sec");
            callMouseEvent(color_choice);
        }, 1);
    }

    // hl green
    if ((key.altKey && key.code === "KeyG") || (key.shiftKey && key.ctrlKey && key.code === "KeyG")) {
        buttonbg = document.getElementById("bgColorButton");

        callMouseEvent(buttonbg);
        setTimeout(function(){
            var color_choice = document.getElementById("docs-material-colorpalette-cell-114");//buttonbg.querySelector('[title="green"]');
            console.log("clickbutton wait 2sec");
            callMouseEvent(color_choice);
        }, 1);
    }

    // hl red
    if ((key.altKey && key.code === "KeyR") || (key.shiftKey && key.ctrlKey && key.code === "KeyR")) {
        buttonbg = document.getElementById("bgColorButton");

        callMouseEvent(buttonbg);
        setTimeout(function(){
            var color_choice = document.getElementById("docs-material-colorpalette-cell-111");//buttonbg.querySelector('[title="red"]');
            console.log("clickbutton wait 2sec");
            callMouseEvent(color_choice);
        }, 1);
    }

    // zoom
    if ((key.altKey && key.code === "KeyZ") || (key.shiftKey && key.ctrlKey && key.code === "KeyZ")) {
        buttonbg = document.getElementById("zoomSelect");

        callMouseEvent(buttonbg);
        setTimeout(function(){

          // var zoom_choice = document.querySelector('[title="red"]');//document.getElementById(":36");//buttonbg.querySelector('[title="red"]');
          var zoom_choice =  Array.from(document.querySelectorAll('div')).find(el => el.textContent === '125%');
            //console.log("clickbutton wait 2sec");
           callMouseEvent(zoom_choice);
        }, 1);
    }


    //---------------Failed attempt to change the font size:

    function dispatchkeyboard(key) {
        if (key.altKey && key.code === "KeyJ") {

            var divfont = document.getElementById("fontSizeSelect");
            console.log(divfont);
            var inputt = divfont.getElementsByTagName('input')[0];
            inputt.select();
            inputt.value = "6";
            console.log('done');

            var ev = document.createEvent('Event');
            ev.initEvent('keypress');
            ev.which = ev.keyCode = 13;
            console.log(ev);
            inputt.dispatchEvent(ev);
        }
    }


}// end of  dispatchkeyboard


//call each mouse event
function callMouseEvent(button){
    triggerMouseEvent (button, "mouseover");
    triggerMouseEvent (button, "mousedown");
    triggerMouseEvent (button, "mouseup");
}

// send mouse even
function triggerMouseEvent (node, eventType) {
    var eventObj = document.createEvent('MouseEvents');
    eventObj.initEvent(eventType, true, true);
    node.dispatchEvent(eventObj);
}


