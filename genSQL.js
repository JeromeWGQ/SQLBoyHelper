// document.onkeydown = function (e) {
//     var keyCode = e.keyCode || e.which || e.charCode;
//     var ctrlKey = e.ctrlKey || e.metaKey;
//     if (ctrlKey && keyCode == 66) {
//         // alert('按住了 CTRL+B');

//     }
//     e.preventDefault();
//     return false;
// }

// function copy(table_name) {
//     navigator && navigator.clipboard && navigator.clipboard.writeText(table_name)
// }

// $("#app > div > div.sidebar-container > div > div.ms-sidebar-main > div > div.router-view > div > div > div.mtd-tabs-content > div > div > div.sql-item-main > div.sql-item-edit > div.sql-item-edit-ds > div > div.table-info > div.key-list").click(function () {
//     let table_name = $("#app > div > div.sidebar-container > div > div.ms-sidebar-main > div > div.router-view > div > div > div.mtd-tabs-content > div > div > div.sql-item-main > div.sql-item-edit > div.sql-item-edit-ds > div > div.table-info > div.key-list > div:nth-child(1) > div > div > div:nth-child(1)").text()
//     // console.log(table_name)
//     copy("select *\nfrom " + table_name + "\nwhere pt_dt='$$yesterday'\n-- and to_date(create_time)='$$yesterday'\norder by rand()\nlimit 100")
// })

// let sql1 = ["-- ", "\nselect *\nfrom ", "\nwhere pt_dt='$$yesterday'\n-- and to_date(create_time)='$$yesterday'\norder by rand()\nlimit 100"]

let buttonScript1 = function () {
    let tableName = this.parentNode.children[0].innerHTML.replace(/<i.*?>/, '').replace(/<\/i>/, '');
    let tableAnnotation = this.parentNode.children[1].innerHTML.replace(/#.*/, '');
    navigator.clipboard.writeText('-- ' + tableAnnotation + ' - ' + tableName + ' - 示例数据\nselect\n*\nfrom ' + tableName + '\nwhere pt_dt=\'$$yesterday\'\n-- and to_date(create_time)=\'$$yesterday\'\norder by rand()\nlimit 100\n');
}

var buttonScriptText1 = buttonScript1.toLocaleString()
var buttonScriptText1 = buttonScriptText1.substring(14, buttonScriptText1.length - 1)

let buttonScript2 = function () {
    let tableName = this.parentNode.children[0].innerHTML.replace(/<i.*?>/, '').replace(/<\/i>/, '');
    let tableAnnotation = this.parentNode.children[1].innerHTML.replace(/#.*/, '');
    navigator.clipboard.writeText('-- ' + tableAnnotation + ' - ' + tableName + ' - 数据量\nselect\npt_dt\n,sum(1)\nfrom ' + tableName + '\n-- select\n-- to_date(create_time)\n-- ,sum(1)\n-- from ' + tableName + '\n-- where pt_dt=\'$$yesterday\'\ngroup by 1\norder by 1 desc\nlimit 30\n');
}

var buttonScriptText2 = buttonScript2.toLocaleString()
var buttonScriptText2 = buttonScriptText2.substring(14, buttonScriptText2.length - 1)

function update() {
    setTimeout(() => {
        $("#app > div > div.sidebar-container > div > div.ms-sidebar-main > div > div.router-view > div > div > div.mtd-tabs-content > div > div > div.sql-item-main > div.sql-item-edit > div.sql-item-edit-ds > div > div.table-info > div.key-list > div:nth-child(n) > div > div").append('<button onclick="' + buttonScriptText1 + '">SQL1</button>').append('<button onclick="' + buttonScriptText2 + '">SQL2</button>')
    }, 1000);
}

// update()

$("#app > div > div.sidebar-container > div > div.ms-sidebar-main > div > div.router-view > div > div > div.mtd-tabs-content > div > div > div.sql-item-main > div.sql-item-edit > div.sql-item-edit-ds > div > div.table-info > div.table-search-input.mtd-input-wrapper.mtd-input-prefix.mtd-input-small > input").keyup(function (event) {
    if (event.which == 13) {
        update()
    }
})

// const source = document.querySelector('.source');

// source.addEventListener('copy', (event) => {
//     const selection = document.getSelection();
//     event.clipboardData.setData('text/plain', selection.toString().toUpperCase());
//     event.preventDefault();
// });
