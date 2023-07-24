
let buttonScript1 = function () {
    let tableName = this.parentNode.children[0].innerHTML.replace(/<i.*?>/, '').replace(/<\/i>/, '');
    let tableAnnotation = this.parentNode.children[1].innerHTML.replace(/#.*/, '');
    navigator.clipboard.writeText('-- ' + tableAnnotation + ' - ' + tableName + ' - 示例数据\nselect\n*\nfrom ' + tableName + '\nwhere pt_dt=\'$$yesterday\'\n-- and to_date(create_time)=\'$$yesterday\'\norder by rand()\nlimit 200\n');
    document.title = tableName.substr(-16);
}

var buttonScriptText1 = buttonScript1.toLocaleString()
var buttonScriptText1 = buttonScriptText1.substring(14, buttonScriptText1.length - 1)

let buttonScript2 = function () {
    let tableName = this.parentNode.children[0].innerHTML.replace(/<i.*?>/, '').replace(/<\/i>/, '');
    let tableAnnotation = this.parentNode.children[1].innerHTML.replace(/#.*/, '');
    navigator.clipboard.writeText('-- ' + tableAnnotation + ' - ' + tableName + ' - 数据量\nselect\npt_dt\n,sum(1)\nfrom ' + tableName + '\n-- select\n-- to_date(create_time)\n-- ,sum(1)\n-- from ' + tableName + '\n-- where pt_dt=\'$$yesterday\'\ngroup by 1\norder by 1 desc\nlimit 100\n');
    document.title = tableName.substr(-16);
}

var buttonScriptText2 = buttonScript2.toLocaleString()
var buttonScriptText2 = buttonScriptText2.substring(14, buttonScriptText2.length - 1)

function update() {
    setTimeout(() => {
        $("#app > div > div.sidebar-container > div > div.ms-sidebar-main > div > div.router-view > div > div > div.mtd-tabs-content > div > div > div.sql-item-main > div.sql-item-edit > div.workspace-left > div > div.table-info > div.key-list > div > div > div").append('<button onclick="' + buttonScriptText1 + '">示例数据</button>').append('<button onclick="' + buttonScriptText2 + '">数据量</button>')
    }, 1000);
}

// 输入框按下回车后，展示按钮
$("#app > div > div.sidebar-container > div > div.ms-sidebar-main > div > div.router-view > div > div > div.mtd-tabs-content > div > div > div.sql-item-main > div.sql-item-edit > div.workspace-left > div > div.table-info > div.table-search-input.mtd-input-wrapper.mtd-input-prefix.mtd-input-small > input").keyup(function (event) {
    if (event.which == 13) {
        update()
    }
})
