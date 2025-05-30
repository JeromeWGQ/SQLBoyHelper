// 定义SQL模板
const SQL_TEMPLATES = {
    sampleData: (tableName, annotation) =>
        `-- ${annotation} - ${tableName} - 示例数据
select
*
from ${tableName}
where pt_dt='$$yesterday'
-- where dt='$$yesterday_compact'
-- and to_date(create_time)='$$yesterday'
order by rand()
limit 200
`,
    dataCount: (tableName, annotation) =>
        `-- ${annotation} - ${tableName} - 数据量
select
pt_dt
-- dt
,count(1)
from ${tableName}
-- select
-- to_date(create_time)
-- ,count(1)
-- from ${tableName}
-- where pt_dt='$$yesterday'
-- where dt='$$yesterday_compact'
group by 1
order by 1 desc
limit 100
`
};

// DOM 选择器
const SELECTORS = {
    tableName: "#app > div > div.sidebar-container > div > div.ms-sidebar-main > div > div > div.router-view > div > div > div.mtd-tabs-content > div > div > div.sql-item-main > div.sql-item-edit > div.workspace-left > div > div.table-info > div.key-list > div > div.table-line.show-column > div > div:nth-child(1)",
    inputBox: "#app > div > div.sidebar-container > div > div.ms-sidebar-main > div > div > div.router-view > div > div > div.mtd-tabs-content > div > div > div.sql-item-main > div.sql-item-edit > div.workspace-left > div > div.table-info > div.table-search-input.mtd-input-wrapper.mtd-input-prefix.mtd-input-small > input"
};

// 按钮处理函数
function handleButtonClick(type) {
    return function() {
        try {
            const tableName = this.parentNode.children[0].innerHTML.replace(/<i.*?>/, '').replace(/<\/i>/, '');
            const tableAnnotation = this.parentNode.children[1].innerHTML.replace(/#.*/, '');
            const sql = SQL_TEMPLATEStype;

            navigator.clipboard.writeText(sql);
            document.title = tableName.substr(-16);
        } catch (error) {
            console.error('按钮点击处理错误:', error);
        }
    }
}

// 更新DOM
function update() {
    setTimeout(() => {
        try {
            const container = $(SELECTORS.tableName);
            container.append(
                $('<button>')
                    .text('示例数据')
                    .on('click', handleButtonClick('sampleData'))
            ).append(
                $('<button>')
                    .text('数据量')
                    .on('click', handleButtonClick('dataCount'))
            );
        } catch (error) {
            console.error('更新DOM错误:', error);
        }
    }, 1000);
}

document.addEventListener('DOMContentLoaded', function() {
    // 输入框按下回车后，展示按钮
    $(SELECTORS.inputBox).keyup(function (event) {  // 修正这里
        if (event.which == 13) {
            console.log('按了回车');
            update();
        }
    });
});

console.log('end');
