
let fillScript = function () {
    elements = document.getElementsByClassName('mtd-input');
    num = document.getElementById('fillInput').value;
    const eInput = new Event('input');
    elements.forEach(element => {
        if (element.placeholder == '长度') {
            element.value = num;
            element.dispatchEvent(eInput);
        }
    });
}

var fillScriptText = fillScript.toLocaleString()
var fillScriptText = fillScriptText.substring(14, fillScriptText.length - 1)

function update() {
    setTimeout(() => {
        $('body > div.app-main > section > section > section > div > div.mtd-loading-nested > div > div:nth-child(4)').append('<input id="fillInput"></input><button onclick="' + fillScriptText + '">fill</button>')
    }, 1000);
}

update()

// $('body > div.app-main > section > section > section > div > div.mtd-loading-nested > div > div:nth-child(2) > div > div:nth-child(1) > div > span.value').click(function () {
//     update()
// })
