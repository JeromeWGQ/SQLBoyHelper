import sys


def print_new_line(sql_type, index, comma):
    output = '    ' + comma + '`' + field_name[index] + '`'
    if sql_type == 1:
        output += ' ' * (max_field_name_length - len(field_name[index]) + 1) + 'COMMENT' + ' \'' + comment[index] + '\''
    print(output)


if __name__ == '__main__':
    field_name = []
    comment = []
    field_type = []
    max_field_name_length = 0
    for line in sys.stdin:
        if line == '\n':
            print('========== 以下为SQL输出 ==========')
            print()
            print('DROP VIEW IF EXISTS `$target.table`;')
            print('CREATE VIEW IF NOT EXISTS `$target.table`')
            print('(')
            print_new_line(1, 0, ' ')
            for i in range(1, len(field_name)):
                print_new_line(1, i, ',')
            print(') COMMENT \'\'')
            print('AS')
            print('select')
            print_new_line(2, 0, ' ')
            for i in range(1, len(field_name)):
                print_new_line(2, i, ',')
            print('from ;')
            sys.exit(0)
        line_splited = line.split('\n')[0].split('\t')
        l1 = line_splited[0]
        l2 = line_splited[1]
        l3 = line_splited[2]
        # 做一点处理
        if l1 == '字段名' or '_update_timestamp' in l1:
            continue
        elif '[P]' in l1:
            l1 = l1.split(' ')[0]
        if l1 == 'id' and l2 == '':
            l2 = '主键'
        # 记录最大长度
        if len(l1) > max_field_name_length:
            max_field_name_length = len(l1)
        # 写入数组
        field_name.append(l1)
        comment.append(l2)
        field_type.append(l3)
