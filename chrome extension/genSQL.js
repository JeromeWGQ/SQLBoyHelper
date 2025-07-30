// 定义SQL模板
const SQL_TEMPLATES = {
    sampleData: (tableName, annotation) =>
        `-- ${annotation} - ${tableName} - 示例数据
select
t.*
from ${tableName} t
where t.pt_dt='$$yesterday'
order by rand()
limit 200
`,
    dataCount: (tableName, annotation) =>
        `-- ${annotation} - ${tableName} - 数据量
select
t.pt_dt
,count(1)
from ${tableName} t
where t.pt_dt >= '$$today{-100d}'
group by 1
order by 1 desc
limit 100
`
};

// DOM 选择器
const SELECTORS = {
//    tableNameNew: "#app > div > div.api-status-container.sql-container-wrapper > div > div > div > div.wrapper.sql-edit-sider-wrapper.collapse-container > div > div > div.main-sider.sider > div.hierarchy-container.meta > div.hierarchy-item.hierarchy-item.tablemeta-container > div.result-container.tablemeta-result > div.all-layout.result > div:nth-child(1) > div > div.table-info > div.table-info-title > span",
    tableNameNew: "#app > div > div.api-status-container.sql-container-wrapper > div > div > div > div.wrapper.sql-edit-sider-wrapper.collapse-container > div > div > div.main-sider.sider > div.hierarchy-container.meta > div.hierarchy-item.hierarchy-item.tablemeta-container > div.result-container.tablemeta-result > div.all-layout.result > div.table-wrapper > div.table > div.table-info > div.table-info-title > span > span",
//    inputBoxNew: "#app > div > div.api-status-container.sql-container-wrapper > div > div > div > div.wrapper.sql-edit-sider-wrapper.collapse-container > div > div > div.main-sider.sider > div.hierarchy-container.meta > div.hierarchy-item.hierarchy-item.tablemeta-container > div.filter-container > div.lower > div.lower-search.mtd-input-wrapper.mtd-input-prefix.mtd-input-suffix > input"
    inputBoxNew: "#app > div > div.api-status-container.sql-container-wrapper > div > div > div > div.wrapper.sql-edit-sider-wrapper.collapse-container > div > div > div.main-sider.sider > div.hierarchy-container.meta > div.hierarchy-item.hierarchy-item.tablemeta-container > div.sqlv2-filter-container > div.lower > div.lower-search.mtd-input-wrapper.mtd-input-prefix.mtd-input-suffix > input"
};

// 按钮处理函数
function handleButtonClick(type) {
    return function() {
        try {
            // 获取表名和表注释
            const tableInfo = document.querySelector("div.table-info");

            // 获取表名
            const tableNameElement = tableInfo.querySelector(".table-info-name");
            const tableNamePrefix = tableNameElement.childNodes[0].textContent.trim(); // "mart_bikedw.app_"
            const tableNameHighlight = tableNameElement.querySelector("span[style*='font-weight: 500']").textContent.trim().split('示例')[0]; // "spock_fault_link"
            const tableName = tableNamePrefix + tableNameHighlight;

            // 获取表注释
            const tableAnnotation = tableInfo.querySelector(".table-info-comment").textContent.trim();

            console.log("表名:", tableName);
            console.log("表注释:", tableAnnotation);

            const sql = SQL_TEMPLATES[type](tableName, tableAnnotation);
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
            const container = $(SELECTORS.tableNameNew);
            if (container.length > 0) {
                container.append(
                    $('<button>')
                        .text('示例数据')
                        .on('click', handleButtonClick('sampleData'))
                ).append(
                    $('<button>')
                        .text('数据量')
                        .on('click', handleButtonClick('dataCount'))
                );
            }
        } catch (error) {
            console.error('更新DOM错误:', error);
        }
    }, 1000);
}

// 初始化函数
function initialize() {
    // 使用事件委托，监听文档上的键盘事件
    $(document).on('keyup', SELECTORS.inputBoxNew, function(event) {
        if (event.which == 13) {
            console.log('按了回车');
            update();
        }
    });
}

// 确保DOM加载完成后执行初始化
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initialize);
} else {
    initialize();
}