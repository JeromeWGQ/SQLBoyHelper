import pandas as pd
import openpyxl
from pypinyin import pinyin, Style
from collections import Counter
import re


def read_excel_range_by_columns(file_path, sheet_name, start_cell, end_cell):
    """
    读取Excel指定区域数据，按纵向（列）顺序存储到数组中

    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名称
        start_cell: 起始单元格 (如 'B2')
        end_cell: 结束单元格 (如 'Y25')

    Returns:
        list: 按列顺序读取的所有数据
    """
    try:
        # 使用pandas读取
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

        # 解析起始和结束位置
        start_col = ord(start_cell[0]) - ord('A')  # B列对应索引1
        start_row = int(start_cell[1:]) - 1  # 第2行对应索引1
        end_col = ord(end_cell[0]) - ord('A')  # Y列对应索引24
        end_row = int(end_cell[1:]) - 1  # 第25行对应索引24

        # 截取指定区域
        region_data = df.iloc[start_row:end_row + 1, start_col:end_col + 1]

        # 按列顺序读取数据到一维数组
        result_array = []
        for col in range(region_data.shape[1]):  # 遍历每一列
            column_data = region_data.iloc[:, col].tolist()
            result_array.extend(column_data)

        # 过滤掉空值（NaN）
        result_array = [str(item) for item in result_array if pd.notna(item)]

        return result_array

    except Exception as e:
        print(f"读取Excel文件出错: {e}")
        return []


def get_char_pinyin(word, char_index=0):
    """
    获取词语指定位置字符的拼音（数字声调形式）

    Args:
        word: 词语字符串
        char_index: 字符索引位置（0表示第一个字，1表示第二个字）

    Returns:
        str: 指定字符的拼音，如果获取失败返回None
    """
    if not word or word == "-" or len(word) <= char_index:
        return None

    try:
        # 获取指定位置的字符
        target_char = word[char_index]

        # 检查是否为汉字
        if '\u4e00' <= target_char <= '\u9fff':
            # 获取拼音（数字声调形式）
            pinyin_result = pinyin(target_char, style=Style.TONE2)
            if pinyin_result and len(pinyin_result[0]) > 0:
                return pinyin_result[0][0]

        return None

    except Exception as e:
        print(f"获取拼音出错 '{word}' 第{char_index + 1}个字: {e}")
        return None


def get_first_char_pinyin(word):
    """获取词语第一个字的拼音"""
    return get_char_pinyin(word, 0)


def get_second_char_pinyin(word):
    """获取词语第二个字的拼音"""
    return get_char_pinyin(word, 1)


def extract_tone_number(pinyin_str):
    """
    从拼音字符串中提取声调数字

    Args:
        pinyin_str: 拼音字符串，如 'hong2'

    Returns:
        int: 声调数字 (1-4)，轻声返回0，无法识别返回None
    """
    if not pinyin_str:
        return None

    # 提取数字
    tone_match = re.search(r'(\d)', pinyin_str)
    if tone_match:
        return int(tone_match.group(1))

    return 0  # 轻声


def adjust_duplicate_pinyin(words, original_pinyin_list):
    """
    调整重复拼音：对于重复的拼音，从第二次出现开始使用第二个字的拼音

    Args:
        words: 词语列表
        original_pinyin_list: 原始拼音列表（第一个字的拼音）

    Returns:
        tuple: (调整后的拼音列表, 调整记录)
    """
    adjusted_pinyin_list = []
    pinyin_seen = {}  # 记录每个拼音第一次出现的位置
    adjustment_records = []  # 记录调整信息

    for i, (word, original_pinyin) in enumerate(zip(words, original_pinyin_list)):
        if original_pinyin is None:
            adjusted_pinyin_list.append(None)
            adjustment_records.append({
                'index': i,
                'word': word,
                'original_pinyin': None,
                'final_pinyin': None,
                'adjusted': False,
                'reason': '无法获取拼音'
            })
            continue

        # 检查是否已经出现过这个拼音
        if original_pinyin in pinyin_seen:
            # 尝试使用第二个字的拼音
            second_char_pinyin = get_second_char_pinyin(word)

            if second_char_pinyin:
                adjusted_pinyin_list.append(second_char_pinyin)
                adjustment_records.append({
                    'index': i,
                    'word': word,
                    'original_pinyin': original_pinyin,
                    'final_pinyin': second_char_pinyin,
                    'adjusted': True,
                    'reason': f'第一个字拼音重复，改用第二个字'
                })
            else:
                # 如果第二个字也无法获取拼音，保持原拼音
                adjusted_pinyin_list.append(original_pinyin)
                adjustment_records.append({
                    'index': i,
                    'word': word,
                    'original_pinyin': original_pinyin,
                    'final_pinyin': original_pinyin,
                    'adjusted': False,
                    'reason': '第二个字无法获取拼音，保持原拼音'
                })
        else:
            # 第一次出现，记录位置并使用原拼音
            pinyin_seen[original_pinyin] = i
            adjusted_pinyin_list.append(original_pinyin)
            adjustment_records.append({
                'index': i,
                'word': word,
                'original_pinyin': original_pinyin,
                'final_pinyin': original_pinyin,
                'adjusted': False,
                'reason': '第一次出现，使用第一个字拼音'
            })

    return adjusted_pinyin_list, adjustment_records


def analyze_pinyin_duplicates(pinyin_list):
    """
    分析拼音重复情况

    Args:
        pinyin_list: 拼音列表

    Returns:
        tuple: (重复标识列表, 统计信息)
    """
    # 过滤掉None值进行统计
    valid_pinyin_list = [p for p in pinyin_list if p is not None]

    # 统计每个拼音的出现次数
    pinyin_counter = Counter(valid_pinyin_list)

    # 生成重复标识列表
    duplicate_flags = []
    for pinyin in pinyin_list:
        if pinyin is None:
            duplicate_flags.append(False)
        elif pinyin_counter[pinyin] > 1:
            duplicate_flags.append(True)  # 有重复
        else:
            duplicate_flags.append(False)  # 无重复

    # 统计信息
    total_count = len(pinyin_list)
    valid_count = len(valid_pinyin_list)
    unique_count = len(pinyin_counter)
    duplicate_count = sum(1 for count in pinyin_counter.values() if count > 1)

    stats = {
        'total_words': total_count,
        'valid_pinyin_count': valid_count,
        'unique_pinyin': unique_count,
        'duplicate_pinyin_types': duplicate_count,
        'duplicate_word_count': sum(duplicate_flags),
        'pinyin_frequency': dict(pinyin_counter)
    }

    return duplicate_flags, stats


def analyze_tone_distribution(pinyin_list):
    """
    分析声调分布情况

    Args:
        pinyin_list: 拼音列表

    Returns:
        dict: 声调统计信息
    """
    tone_counter = Counter()

    for pinyin in pinyin_list:
        if pinyin:
            tone = extract_tone_number(pinyin)
            if tone is not None:
                tone_counter[tone] += 1

    # 声调名称映射
    tone_names = {
        1: '一声(阴平)',
        2: '二声(阳平)',
        3: '三声(上声)',
        4: '四声(去声)',
        0: '轻声'
    }

    tone_stats = {}
    total_with_tone = sum(tone_counter.values())

    for tone_num in [1, 2, 3, 4, 0]:
        count = tone_counter.get(tone_num, 0)
        percentage = (count / total_with_tone * 100) if total_with_tone > 0 else 0
        tone_stats[tone_num] = {
            'name': tone_names[tone_num],
            'count': count,
            'percentage': percentage
        }

    return tone_stats


def print_adjustment_summary(adjustment_records):
    """
    打印调整摘要信息
    """
    print("\n" + "=" * 80)
    print("拼音调整摘要")
    print("=" * 80)

    adjusted_count = sum(1 for record in adjustment_records if record['adjusted'])
    total_count = len(adjustment_records)

    print(f"\n📊 调整统计:")
    print(f"   总词数: {total_count}")
    print(f"   调整词数: {adjusted_count}")
    print(f"   调整比例: {adjusted_count / total_count * 100:.1f}%")

    # 显示调整的词语
    if adjusted_count > 0:
        print(f"\n🔄 调整详情 (前20个):")
        print(f"{'序号':<4} {'词语':<8} {'原拼音':<8} {'新拼音':<8} {'原因'}")
        print("-" * 60)

        adjusted_records = [r for r in adjustment_records if r['adjusted']]
        for i, record in enumerate(adjusted_records[:20]):
            print(
                f"{i + 1:<4} {record['word']:<8} {record['original_pinyin']:<8} {record['final_pinyin']:<8} {record['reason']}")

        if len(adjusted_records) > 20:
            print(f"   ... 还有{len(adjusted_records) - 20}个调整")


def print_detailed_results(words, pinyin_list, duplicate_flags, stats, tone_stats, title="分析结果"):
    """
    打印详细的分析结果
    """
    print("\n" + "=" * 80)
    print(f"{title}")
    print("=" * 80)

    # 基本统计
    print(f"\n📊 基本统计:")
    print(f"   总词数: {stats['total_words']}")
    print(f"   有效拼音数: {stats['valid_pinyin_count']}")
    print(f"   唯一拼音数: {stats['unique_pinyin']}")
    print(f"   有重复的拼音种类数: {stats['duplicate_pinyin_types']}")
    print(f"   有重复拼音的词数: {stats['duplicate_word_count']}")

    # 声调分布统计
    print(f"\n🎵 声调分布统计:")
    for tone_num in [1, 2, 3, 4, 0]:
        tone_info = tone_stats[tone_num]
        print(f"   {tone_info['name']:<12}: {tone_info['count']:4d}次 ({tone_info['percentage']:5.1f}%)")

    # 拼音频率统计（按频率排序）
    print(f"\n🔤 拼音频率统计 (前15个):")
    sorted_freq = sorted(stats['pinyin_frequency'].items(), key=lambda x: x[1], reverse=True)
    for i, (pinyin, count) in enumerate(sorted_freq[:15]):
        tone = extract_tone_number(pinyin)
        tone_name = tone_stats.get(tone, {}).get('name', '未知') if tone is not None else '未知'
        print(f"   {i + 1:2d}. {pinyin:8s} - {count:3d}次 ({tone_name})")

    # 显示重复拼音的详细信息
    duplicate_groups = {}
    for word, pinyin, is_duplicate in zip(words, pinyin_list, duplicate_flags):
        if is_duplicate and pinyin:
            if pinyin not in duplicate_groups:
                duplicate_groups[pinyin] = []
            duplicate_groups[pinyin].append(word)

    if duplicate_groups:
        print(f"\n🔄 重复拼音详情 (前10个):")
        for i, (pinyin, word_list) in enumerate(sorted(duplicate_groups.items())[:10]):
            print(f"   {pinyin:8s} ({len(word_list)}个): {', '.join(word_list[:8])}")
            if len(word_list) > 8:
                print(f"            ... 还有{len(word_list) - 8}个")


def save_results_to_file(words, original_pinyin_list, adjusted_pinyin_list,
                         original_duplicate_flags, adjusted_duplicate_flags,
                         original_stats, adjusted_stats, original_tone_stats, adjusted_tone_stats,
                         adjustment_records, filename="comprehensive_analysis_results.txt"):
    """
    将完整结果保存到文件
    """
    try:
        with open(filename, 'w', encoding='utf-8') as f:
            f.write("Excel词语拼音综合分析结果\n")
            f.write("=" * 60 + "\n\n")

            # 原始分析结果
            f.write("【原始分析结果】\n")
            f.write(f"总词数: {original_stats['total_words']}\n")
            f.write(f"有效拼音数: {original_stats['valid_pinyin_count']}\n")
            f.write(f"唯一拼音数: {original_stats['unique_pinyin']}\n")
            f.write(f"有重复的拼音种类数: {original_stats['duplicate_pinyin_types']}\n")
            f.write(f"有重复拼音的词数: {original_stats['duplicate_word_count']}\n\n")

            # 调整后分析结果
            f.write("【调整后分析结果】\n")
            f.write(f"总词数: {adjusted_stats['total_words']}\n")
            f.write(f"有效拼音数: {adjusted_stats['valid_pinyin_count']}\n")
            f.write(f"唯一拼音数: {adjusted_stats['unique_pinyin']}\n")
            f.write(f"有重复的拼音种类数: {adjusted_stats['duplicate_pinyin_types']}\n")
            f.write(f"有重复拼音的词数: {adjusted_stats['duplicate_word_count']}\n\n")

            # 声调统计对比
            f.write("【声调分布对比】\n")
            f.write("声调\t\t原始\t\t调整后\n")
            for tone_num in [1, 2, 3, 4, 0]:
                orig_info = original_tone_stats[tone_num]
                adj_info = adjusted_tone_stats[tone_num]
                f.write(
                    f"{orig_info['name']}\t{orig_info['count']}次({orig_info['percentage']:.1f}%)\t{adj_info['count']}次({adj_info['percentage']:.1f}%)\n")
            f.write("\n")

            # 调整记录
            adjusted_count = sum(1 for r in adjustment_records if r['adjusted'])
            f.write(f"【调整记录】(共{adjusted_count}个调整)\n")
            f.write("序号\t词语\t原拼音\t新拼音\t调整原因\n")
            adj_idx = 1
            for record in adjustment_records:
                if record['adjusted']:
                    f.write(
                        f"{adj_idx}\t{record['word']}\t{record['original_pinyin']}\t{record['final_pinyin']}\t{record['reason']}\n")
                    adj_idx += 1
            f.write("\n")

            # 详细数据
            f.write("【详细数据】\n")
            f.write("序号\t词语\t原拼音\t调整后拼音\t原始重复\t调整后重复\n")
            for i, (word, orig_pinyin, adj_pinyin, orig_dup, adj_dup) in enumerate(
                    zip(words, original_pinyin_list, adjusted_pinyin_list, original_duplicate_flags,
                        adjusted_duplicate_flags)):
                orig_pinyin_str = orig_pinyin if orig_pinyin else "N/A"
                adj_pinyin_str = adj_pinyin if adj_pinyin else "N/A"
                orig_dup_str = "是" if orig_dup else "否"
                adj_dup_str = "是" if adj_dup else "否"
                f.write(f"{i + 1}\t{word}\t{orig_pinyin_str}\t{adj_pinyin_str}\t{orig_dup_str}\t{adj_dup_str}\n")

        print(f"\n💾 完整结果已保存到文件: {filename}")

    except Exception as e:
        print(f"保存文件出错: {e}")


def main():
    """
    主函数
    """
    # 文件路径和参数设置
    file_path = 'ci-test.xlsx'  # 请修改为你的文件路径
    sheet_name = 'ci-test'

    print("🚀 开始读取Excel文件...")

    # 读取Excel数据
    words = read_excel_range_by_columns(file_path, sheet_name, 'B2', 'Y25')

    if not words:
        print("❌ 没有读取到数据，请检查文件路径和工作表名称")
        return

    print(f"✅ 成功读取到 {len(words)} 个词语")

    # 过滤掉"-"并获取拼音
    print("\n🔤 正在获取原始拼音...")
    filtered_words = []
    original_pinyin_list = []

    for word in words:
        if word != "-":  # 排除"-"
            filtered_words.append(word)
            pinyin = get_first_char_pinyin(word)
            original_pinyin_list.append(pinyin)

    # 过滤掉无法获取拼音的词
    valid_data = [(w, p) for w, p in zip(filtered_words, original_pinyin_list) if p is not None]

    if not valid_data:
        print("❌ 没有找到有效的汉字词语")
        return

    final_words, final_original_pinyin_list = zip(*valid_data)
    final_words = list(final_words)
    final_original_pinyin_list = list(final_original_pinyin_list)

    print(f"✅ 成功获取 {len(final_original_pinyin_list)} 个词语的拼音")

    # 分析原始重复情况
    print("\n🔍 正在分析原始拼音重复情况...")
    original_duplicate_flags, original_stats = analyze_pinyin_duplicates(final_original_pinyin_list)
    original_tone_stats = analyze_tone_distribution(final_original_pinyin_list)

    # 打印原始结果
    print_detailed_results(final_words, final_original_pinyin_list, original_duplicate_flags,
                           original_stats, original_tone_stats, "原始分析结果")

    # 调整重复拼音
    print("\n🔧 正在调整重复拼音...")
    adjusted_pinyin_list, adjustment_records = adjust_duplicate_pinyin(final_words, final_original_pinyin_list)

    # 打印调整摘要
    print_adjustment_summary(adjustment_records)

    # 分析调整后的重复情况
    print("\n🔍 正在分析调整后拼音重复情况...")
    adjusted_duplicate_flags, adjusted_stats = analyze_pinyin_duplicates(adjusted_pinyin_list)
    adjusted_tone_stats = analyze_tone_distribution(adjusted_pinyin_list)

    # 打印调整后结果
    print_detailed_results(final_words, adjusted_pinyin_list, adjusted_duplicate_flags,
                           adjusted_stats, adjusted_tone_stats, "调整后分析结果")

    # 对比分析
    print("\n" + "=" * 80)
    print("对比分析")
    print("=" * 80)

    print(f"\n📈 重复情况改善:")
    original_duplicate_count = original_stats['duplicate_word_count']
    adjusted_duplicate_count = adjusted_stats['duplicate_word_count']
    improvement = original_duplicate_count - adjusted_duplicate_count
    improvement_rate = (improvement / original_duplicate_count * 100) if original_duplicate_count > 0 else 0

    print(f"   原始重复词数: {original_duplicate_count}")
    print(f"   调整后重复词数: {adjusted_duplicate_count}")
    print(f"   减少重复词数: {improvement}")
    print(f"   改善率: {improvement_rate:.1f}%")

    print(f"\n🎵 声调分布变化:")
    for tone_num in [1, 2, 3, 4, 0]:
        orig_count = original_tone_stats[tone_num]['count']
        adj_count = adjusted_tone_stats[tone_num]['count']
        change = adj_count - orig_count
        change_str = f"({change:+d})" if change != 0 else ""
        print(f"   {original_tone_stats[tone_num]['name']}: {orig_count} → {adj_count} {change_str}")

    # 保存完整结果到文件
    save_results_to_file(final_words, final_original_pinyin_list, adjusted_pinyin_list,
                         original_duplicate_flags, adjusted_duplicate_flags,
                         original_stats, adjusted_stats, original_tone_stats, adjusted_tone_stats,
                         adjustment_records)

    # 返回结果
    return {
        'words': final_words,
        'original_pinyin_list': final_original_pinyin_list,
        'adjusted_pinyin_list': adjusted_pinyin_list,
        'original_duplicate_flags': original_duplicate_flags,
        'adjusted_duplicate_flags': adjusted_duplicate_flags,
        'original_stats': original_stats,
        'adjusted_stats': adjusted_stats,
        'original_tone_stats': original_tone_stats,
        'adjusted_tone_stats': adjusted_tone_stats,
        'adjustment_records': adjustment_records
    }


if __name__ == "__main__":
    # 安装依赖提示
    print("📦 请确保已安装依赖: pip install pandas openpyxl pypinyin")
    print()

    # 运行脚本
    results = main()

    if results:
        print(f"\n🎉 综合分析完成！")
        print(f"   - 原始拼音数组长度: {len(results['original_pinyin_list'])}")
        print(f"   - 调整后拼音数组长度: {len(results['adjusted_pinyin_list'])}")
        print(f"   - 调整词语数量: {sum(1 for r in results['adjustment_records'] if r['adjusted'])}")
        print(
            f"   - 重复改善: {results['original_stats']['duplicate_word_count']} → {results['adjusted_stats']['duplicate_word_count']}")