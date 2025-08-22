import pandas as pd
from pypinyin import pinyin, Style
from collections import Counter
import re
import csv


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

        # 过滤：仅保留纯中文的值
        filtered_array = []
        for item in result_array:
            if pd.notna(item):
                item_str = str(item).strip()
                # 检查是否为纯中文
                if re.match(r'^[\u4e00-\u9fff]+$', item_str):
                    # 验证字数（必须是1个字或2个字）
                    if len(item_str) == 1 or len(item_str) == 2:
                        filtered_array.append(item_str)
                    else:
                        raise ValueError(f"❌ 发现不符合规则的纯中文词语：'{item_str}'（长度：{len(item_str)}字）\n"
                                       f"   规则要求：纯中文词语必须是1个字或2个字")

        return filtered_array

    except Exception as e:
        print(f"读取Excel文件出错: {e}")
        return []


def check_single_char_pronunciation_duplicates(words):
    """
    检查所有单字的读音是否有重复

    Args:
        words: 词语列表

    Returns:
        bool: True表示无重复，False表示有重复
    """
    print("\n🔍 开始检查单字读音重复情况...")

    # 提取所有单字
    single_chars = [word for word in words if len(word) == 1]

    if not single_chars:
        print("⚠️  没有发现单字，跳过单字读音检查")
        return True

    print(f"📝 发现 {len(single_chars)} 个单字")

    # 获取每个单字的读音（带声调）
    char_pronunciations = {}
    pronunciation_chars = {}  # 用于记录每个读音对应的字符

    for char in single_chars:
        # 获取拼音（带声调数字）
        pronunciation = pinyin(char, style=Style.TONE3)[0][0]
        char_pronunciations[char] = pronunciation

        # 检查读音是否已存在
        if pronunciation in pronunciation_chars:
            # 发现重复读音
            print(f"\n❌ 发现单字读音重复！")
            print(f"   读音：{pronunciation}")
            print(f"   对应字符：'{pronunciation_chars[pronunciation]}' 和 '{char}'")
            print(f"   程序退出")
            return False
        else:
            pronunciation_chars[pronunciation] = char

    print(f"✅ 单字读音检查通过，无重复！")

    return True


def build_pronunciation_collection(words):
    """
    构建所有单字+双字第一个字的读音集合数组（保留重复值）

    Args:
        words: 词语列表

    Returns:
        list: 读音集合数组，保留重复值
    """
    print("🔧 构建读音集合数组...")

    pronunciation_collection = []

    for word in words:
        if len(word) == 1:
            # 单字：使用上下文推断的准确读音
            pronunciation = get_accurate_pronunciation_from_context(word, words)
            pronunciation_collection.append(pronunciation)
        elif len(word) == 2:
            # 双字：获取第一个字的上下文推断读音
            first_char = word[0]
            pronunciation = get_accurate_pronunciation_from_context(first_char, words)
            pronunciation_collection.append(pronunciation)

    return pronunciation_collection


def determine_double_char_pronunciation(words, pronunciation_collection):
    """
    确定双字的读音规则
    如果双字的首字读音在集合中是独一无二的，读首字；否则读原始词

    Args:
        words: 词语列表
        pronunciation_collection: 读音集合数组（包含重复值）

    Returns:
        dict: 词语到读音的映射字典
    """
    print("\n🎯 开始确定双字的读音规则...")

    # 统计读音出现次数
    pronunciation_counter = Counter(pronunciation_collection)

    # 创建词语到读音的映射
    word_pronunciation_mapping = {}

    for word in words:
        if len(word) == 1:
            # 单字：直接使用其读音
            pronunciation = pinyin(word, style=Style.TONE3)[0][0]
            word_pronunciation_mapping[word] = pronunciation
            print(f"   单字 '{word}' → {pronunciation}")

        elif len(word) == 2:
            # 双字：根据规则确定读音
            first_char = word[0]
            first_char_pronunciation = pinyin(first_char, style=Style.TONE3)[0][0]

            # 检查首字读音是否独一无二
            if pronunciation_counter[first_char_pronunciation] == 1:
                # 首字读音独一无二，读首字
                word_pronunciation_mapping[word] = first_char_pronunciation
                print(f"   双字 '{word}' → {first_char_pronunciation} (读首字 '{first_char}'，因为读音独一无二)")
            else:
                # 首字读音不独一无二，读原始词
                full_word_pronunciation = ''.join([pinyin(char, style=Style.TONE3)[0][0] for char in word])
                word_pronunciation_mapping[word] = full_word_pronunciation
                print(f"   双字 '{word}' → {full_word_pronunciation} (读原始词，因为首字读音 '{first_char_pronunciation}' 重复 {pronunciation_counter[first_char_pronunciation]} 次)")

    print(f"\n✅ 双字读音规则确定完成！")
    print(f"📋 最终词语读音映射表：")
    for word, pronunciation in word_pronunciation_mapping.items():
        print(f"   '{word}' → {pronunciation}")

    # 检查最终读音是否有重复
    final_pronunciations = list(word_pronunciation_mapping.values())
    final_pronunciation_counter = Counter(final_pronunciations)
    duplicates = [p for p, count in final_pronunciation_counter.items() if count > 1]

    if duplicates:
        print(f"\n⚠️  警告：最终读音中发现重复！")
        for dup_pronunciation in duplicates:
            dup_words = [word for word, pronunciation in word_pronunciation_mapping.items()
                        if pronunciation == dup_pronunciation]
            print(f"   读音 '{dup_pronunciation}' 对应词语：{dup_words}")
    else:
        print(f"\n✅ 最终读音检查通过，无重复！")

    return word_pronunciation_mapping


def create_pronunciation_mapping_with_chars(words, pronunciation_collection):
    """
    创建词语到读音的映射，读音用汉字表示

    Args:
        words: 词语列表
        pronunciation_collection: 读音集合数组（包含重复值）

    Returns:
        tuple: (word_pronunciation_mapping, detailed_results)
    """
    print("🎯 确定词语读音规则...")

    # 统计读音出现次数
    pronunciation_counter = Counter(pronunciation_collection)

    # 创建词语到读音的映射和详细结果
    word_pronunciation_mapping = {}
    detailed_results = []

    # 统计变量
    single_char_count = 0
    double_char_read_first = 0
    double_char_read_full = 0

    for word in words:
        if len(word) == 1:
            # 单字：读音就是单字本身
            pronunciation_char = word
            # 使用上下文推断的准确读音
            pronunciation_pinyin = get_accurate_pronunciation_from_context(word, words)
            first_char_count = pronunciation_counter[pronunciation_pinyin]

            # 找到与该单字读音相同的所有词汇
            same_pronunciation_words = []
            for w in words:
                if len(w) == 1:
                    if get_accurate_pronunciation_from_context(w, words) == pronunciation_pinyin:
                        same_pronunciation_words.append(w)
                elif len(w) == 2:
                    first_char_pinyin = get_accurate_pronunciation_from_context(w[0], words)
                    if first_char_pinyin == pronunciation_pinyin:
                        same_pronunciation_words.append(w)

            word_pronunciation_mapping[word] = pronunciation_char
            detailed_results.append({
                '原始词': word,
                '读音（汉字）': pronunciation_char,
                '读音（拼音）': pronunciation_pinyin,
                '首字音重复数量': first_char_count,
                '首字音重复词汇列表': ', '.join(same_pronunciation_words)
            })
            single_char_count += 1

        elif len(word) == 2:
            # 双字：根据规则确定读音
            first_char = word[0]
            first_char_pronunciation = get_accurate_pronunciation_from_context(first_char, words)
            first_char_count = pronunciation_counter[first_char_pronunciation]

            # 找到与该双字首字读音相同的所有词汇
            same_pronunciation_words = []
            for w in words:
                if len(w) == 1:
                    if pinyin(w, style=Style.TONE3)[0][0] == first_char_pronunciation:
                        same_pronunciation_words.append(w)
                elif len(w) == 2:
                    first_char_pinyin = pinyin(w[0], style=Style.TONE3)[0][0]
                    if first_char_pinyin == first_char_pronunciation:
                        same_pronunciation_words.append(w)

            # 检查首字读音是否独一无二
            if pronunciation_counter[first_char_pronunciation] == 1:
                # 首字读音独一无二，读首字
                pronunciation_char = first_char
                pronunciation_pinyin = first_char_pronunciation
                word_pronunciation_mapping[word] = pronunciation_char
                double_char_read_first += 1
            else:
                # 首字读音不独一无二，读原始词
                pronunciation_char = word
                pronunciation_pinyin = ''.join([pinyin(char, style=Style.TONE3)[0][0] for char in word])
                word_pronunciation_mapping[word] = pronunciation_char
                double_char_read_full += 1

            detailed_results.append({
                '原始词': word,
                '读音（汉字）': pronunciation_char,
                '读音（拼音）': pronunciation_pinyin,
                '首字音重复数量': first_char_count,
                '首字音重复词汇列表': ', '.join(same_pronunciation_words)
            })

    # 检查最终读音是否有重复
    final_pronunciations = list(word_pronunciation_mapping.values())
    final_pronunciation_counter = Counter(final_pronunciations)
    duplicates = [p for p, count in final_pronunciation_counter.items() if count > 1]

    if duplicates:
        print(f"⚠️  警告：最终读音中发现重复！")
        for dup_pronunciation in duplicates:
            dup_words = [word for word, pronunciation in word_pronunciation_mapping.items()
                        if pronunciation == dup_pronunciation]
            print(f"   读音 '{dup_pronunciation}' 对应词语：{dup_words}")
    else:
        print(f"✅ 最终读音检查通过，无重复！")

    return word_pronunciation_mapping, detailed_results, {
        'single_char_count': single_char_count,
        'double_char_read_first': double_char_read_first,
        'double_char_read_full': double_char_read_full
    }


def save_results_to_csv(detailed_results, filename='词语读音映射结果.csv'):
    """
    将结果保存到CSV文件

    Args:
        detailed_results: 详细结果列表
        filename: 输出文件名
    """
    print(f"\n💾 开始保存结果到CSV文件：{filename}")

    try:
        with open(filename, 'w', newline='', encoding='utf-8-sig') as csvfile:
            fieldnames = ['原始词', '读音（汉字）', '读音（拼音）', '首字音重复数量', '首字音重复词汇列表']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

            # 写入表头
            writer.writeheader()

            # 写入数据
            for result in detailed_results:
                writer.writerow(result)

        print(f"✅ 成功保存 {len(detailed_results)} 条记录到 {filename}")

    except Exception as e:
        print(f"❌ 保存CSV文件时出错: {e}")


def get_accurate_pronunciation_from_context(char, words):
    """
    通过词组上下文推断字符的准确读音

    Args:
        char: 要推断读音的字符
        words: 词语列表

    Returns:
        str: 推断出的最准确读音
    """
    # 找到包含该字符的所有双字词组
    containing_words = [word for word in words if len(word) == 2 and word[0] == char]

    if not containing_words:
        # 如果没有双字词组，使用单字的默认读音
        return pinyin(char, style=Style.TONE3)[0][0]

    # 获取该字符的所有可能读音
    all_pronunciations = pinyin(char, style=Style.TONE3, heteronym=True)[0]

    if len(all_pronunciations) <= 1:
        # 不是多音字，直接返回
        return all_pronunciations[0]

    # 统计每个读音在词组中的使用频率
    pronunciation_votes = {}

    for word in containing_words:
        # 获取整个词组的读音
        word_pronunciations = pinyin(word, style=Style.TONE3)
        first_char_pronunciation = word_pronunciations[0][0]

        # 验证这个读音是否在该字符的可能读音列表中
        if first_char_pronunciation in all_pronunciations:
            pronunciation_votes[first_char_pronunciation] = pronunciation_votes.get(first_char_pronunciation, 0) + 1

    if pronunciation_votes:
        # 返回得票最多的读音
        most_common_pronunciation = max(pronunciation_votes.items(), key=lambda x: x[1])[0]
        return most_common_pronunciation
    else:
        # 如果没有有效的投票，返回默认读音
        return pinyin(char, style=Style.TONE3)[0][0]


def check_polyphonic_characters(words):
    """
    检查多音字情况，通过词组上下文推断准确读音

    Args:
        words: 词语列表

    Returns:
        dict: 多音字信息
    """
    print("🔍 检查多音字情况...")

    # 收集所有需要检查的字符（单字 + 双字的首字）
    chars_to_check = set()

    for word in words:
        if len(word) == 1:
            chars_to_check.add(word)
        elif len(word) == 2:
            chars_to_check.add(word[0])

    polyphonic_info = {}

    for char in chars_to_check:
        # 获取所有可能的读音
        all_pronunciations = pinyin(char, style=Style.TONE3, heteronym=True)[0]

        if len(all_pronunciations) > 1:
            # 这是一个多音字
            # 通过词组上下文推断准确读音
            context_pronunciation = get_accurate_pronunciation_from_context(char, words)
            # 获取默认读音（不考虑上下文）
            default_pronunciation = pinyin(char, style=Style.TONE3)[0][0]

            # 找到包含该字符的词组
            containing_words = [word for word in words if len(word) == 2 and word[0] == char]

            polyphonic_info[char] = {
                'all_pronunciations': all_pronunciations,
                'default_pronunciation': default_pronunciation,
                'context_pronunciation': context_pronunciation,
                'containing_words': containing_words,
                'is_context_different': context_pronunciation != default_pronunciation
            }

    return polyphonic_info


def print_processing_logic():
    """
    打印联想词处理逻辑说明
    """
    print("\n" + "="*60)
    print("📖 联想词读音处理逻辑说明")
    print("="*60)
    print("目标：为每个词语分配唯一的读音，确保无重复")
    print()
    print("处理步骤：")
    print("1️⃣  数据预处理：仅保留纯中文1-2字词语")
    print("2️⃣  单字读音检查：确保所有单字读音无重复")
    print("3️⃣  读音集合构建：收集所有单字+双字首字的读音")
    print("4️⃣  双字读音规则：")
    print("    • 如果双字首字读音独一无二 → 读首字")
    print("    • 如果双字首字读音有重复 → 读完整词语")
    print("5️⃣  最终验证：确保所有词语的读音无重复")
    print("6️⃣  结果输出：生成CSV文件包含详细映射信息")
    print("="*60)




def print_statistics(words, stats, detailed_results):
    """
    打印统计信息
    """
    print("\n" + "="*50)
    print("📊 处理结果统计")
    print("="*50)

    # 基础统计
    total_words = len(words)
    single_chars = [w for w in words if len(w) == 1]
    double_chars = [w for w in words if len(w) == 2]

    print(f"📝 词语总数：{total_words}")
    print(f"   • 单字：{len(single_chars)} 个")
    print(f"   • 双字：{len(double_chars)} 个")
    print()

    # 读音策略统计
    print("🎯 读音策略分布：")
    print(f"   • 单字（读自身）：{stats['single_char_count']} 个")
    print(f"   • 双字（读首字）：{stats['double_char_read_first']} 个")
    print(f"   • 双字（读完整）：{stats['double_char_read_full']} 个")
    print()

    # 读音长度统计
    pronunciation_lengths = {}
    for result in detailed_results:
        pronunciation = result['读音（汉字）']
        length = len(pronunciation)
        pronunciation_lengths[length] = pronunciation_lengths.get(length, 0) + 1

    print("📏 读音长度分布：")
    for length in sorted(pronunciation_lengths.keys()):
        print(f"   • {length}字读音：{pronunciation_lengths[length]} 个")
    print()

    # 首字音重复情况统计
    repeat_counts = {}
    for result in detailed_results:
        count = result['首字音重复数量']
        repeat_counts[count] = repeat_counts.get(count, 0) + 1

    print("🔄 首字音重复情况：")
    for count in sorted(repeat_counts.keys()):
        if count == 1:
            print(f"   • 独一无二：{repeat_counts[count]} 个词语")
        else:
            print(f"   • 重复{count}次：{repeat_counts[count]} 个词语")
    print()

    # 多音字统计
    polyphonic_info = check_polyphonic_characters(words)

    if polyphonic_info:
        # 筛选出读音发生变化的多音字
        changed_polyphonic = {char: info for char, info in polyphonic_info.items() if info['is_context_different']}

        print("🎵 多音字情况：")
        print(f"   发现 {len(polyphonic_info)} 个多音字，其中 {len(changed_polyphonic)} 个通过上下文推断改变了读音")

        if changed_polyphonic:
            print()
            print("📋 上下文推断改变读音的多音字：")
            for char, info in sorted(changed_polyphonic.items()):
                all_sounds = ' / '.join(info['all_pronunciations'])
                default_sound = info['default_pronunciation']
                context_sound = info['context_pronunciation']
                containing_words = ', '.join(info['containing_words'])

                print(f"   '{char}' → 所有读音: [{all_sounds}]")
                print(f"        默认读音: {default_sound} | 上下文推断: {context_sound} ⭐")
                print(f"        参考词组: {containing_words}")
            print()
            print("   ⭐ 表示通过词组上下文推断出的读音与默认读音不同")
        else:
            print("   所有多音字的上下文推断读音与默认读音一致")
    else:
        print("🎵 多音字情况：")
        print("   未发现多音字")

    print("="*50)


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

    print(f"✅ 成功读取到 {len(words)} 个有效词语（仅保留纯中文1-2字词语）")

    # 打印处理逻辑说明
    print_processing_logic()

    # 步骤1：检查所有单字的读音是否有重复
    if not check_single_char_pronunciation_duplicates(words):
        return  # 如果有重复，退出程序

    # 步骤2：构建所有单字+双字第一个字的读音集合数组（保留重复值）
    pronunciation_collection = build_pronunciation_collection(words)

    # 步骤3：确定双字的读音规则（用汉字表示读音）
    word_pronunciation_mapping, detailed_results, stats = create_pronunciation_mapping_with_chars(words, pronunciation_collection)

    # 步骤4：保存结果到CSV文件
    save_results_to_csv(detailed_results)

    # 步骤5：打印统计信息
    print_statistics(words, stats, detailed_results)


if __name__ == "__main__":
    main()
