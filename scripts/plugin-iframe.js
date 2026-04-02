var text = "";
(function load() {
    if (window.captureEvents) {
        document.captureEvents(Event.MOUSEUP);
        document.captureEvents(Event.KEYDOWN);
        document.captureEvents(Event.MOUSEDOWN);
        document.captureEvents(Event.TOUCHEND);
        document.onmouseup = getSelectedText;
        document.onmousedown = clearSelection;
        document.ontouchend = getSelectedText;
    } else {
        document.onmouseup = getSelectedText;
        document.onmousedown = clearSelection;
        document.ontouchend = getSelectedText;
    }

})();

function getSelectedText() {
    console.log("getSelectedText");
    if (window.getSelection) {
        text = "" + window.getSelection();
    } else if (document.getSelection) {
        text = "" + document.getSelection();
    } else if (document.selection) {
        text = "" + document.selection.createRange().text;
    } 
    if (text == "" || text == null) {
        return;
    }
    console.log("text["+text+"]");
    var libras = parent.document.getElementById("rybenaDiv");
    if(libras != null){
        if (libras.style.visibility != "hidden") {
            if (text != " " && text != "  ") {
                parent.translate(text);
            }
        } 
    }
}
function clearSelection() {
    var sel;
    if ((sel = document.selection) && sel.empty) {
        sel.empty();
    } else {
        if (window.getSelection) {
            window.getSelection().removeAllRanges();
        }
        var activeEl = document.activeElement;
        if (activeEl) {
            var tagName = activeEl.nodeName.toLowerCase();
            if (tagName == "textarea"
                    || (tagName == "input" && activeEl.type == "text")) {
                activeEl.selectionStart = activeEl.selectionEnd;
            }
        }
    }
}