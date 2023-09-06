# 脚本调用示例：
# gen_drop_partition.py topic_bike_repair_dispatch_detail_inc p20210721 p20210910
# -参数1：表名 -参数2：要删除的第一个分区 -参数3：要删除的最后一个分区

import sys
import datetime


def get_output_str(table_name1, date):
    return 'alter table `%s` drop partition `p%s`;' % (table_name1, date.strftime('%Y%m%d'))


if __name__ == '__main__':
    args = sys.argv[1:]
    table_name = args[0]
    startdate = datetime.datetime.strptime(args[1][1:], '%Y%m%d')
    enddate = datetime.datetime.strptime(args[2][1:], '%Y%m%d')
    now = datetime.datetime.now()
    if startdate > enddate or startdate > now or enddate > now:
        print('参数有误')
        sys.exit(0)
    print('========== 复制以下输出至[https://data.sankuai.com/wanxiang#/olap/data/bikedw/sql]执行 ==========')
    print(get_output_str(table_name, startdate))
    while startdate != enddate:
        startdate = startdate + datetime.timedelta(days=1)
        print(get_output_str(table_name, startdate))
