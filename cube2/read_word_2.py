import pandas as pd
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

        # 过滤掉空值（NaN）和"-"
        result_array = [str(item) for item in result_array if pd.notna(item) and str(item) != "-"]

        return result_array

    except Exception as e:
        print(f"读取Excel文件出错: {e}")
        return []


def classify_words(words):
    """
    对词汇进行分类：正常汉字词汇、英文词汇、单字词汇、其他特殊词汇

    Args:
        words: 词语列表

    Returns:
        dict: 分类结果
    """
    classification = {
        'normal_chinese': [],  # 正常汉字词汇（多字）
        'single_chinese': [],  # 单个汉字词汇
        'english_words': [],  # 英文词汇
        'mixed_words': [],  # 中英混合词汇
        'special_words': []  # 其他特殊词汇（数字、符号等）
    }

    for word in words:
        if len(word) == 1:
            # 单字词汇
            if '\u4e00' <= word <= '\u9fff':
                classification['single_chinese'].append(word)
            elif word.isalpha() and word.isascii():
                classification['english_words'].append(word)
            else:
                classification['special_words'].append(word)

        elif len(word) > 1:
            # 多字词汇
            chinese_chars = sum(1 for char in word if '\u4e00' <= char <= '\u9fff')
            english_chars = sum(1 for char in word if char.isalpha() and char.isascii())
            other_chars = len(word) - chinese_chars - english_chars

            if chinese_chars == len(word):
                # 纯汉字词汇
                classification['normal_chinese'].append(word)
            elif english_chars == len(word):
                # 纯英文词汇
                classification['english_words'].append(word)
            elif chinese_chars > 0 and english_chars > 0:
                # 中英混合词汇
                classification['mixed_words'].append(word)
            else:
                # 其他特殊词汇
                classification['special_words'].append(word)

    return classification


def get_char_pinyin(word, char_index=0):
    """
    获取词语指定位置字符的拼音（数字声调形式）

    Args:
        word: 词语字符串
        char_index: 字符索引位置（0表示第一个字，1表示第二个字）

    Returns:
        str: 指定字符的拼音，如果获取失败返回None
    """
    if not word or len(word) <= char_index:
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


def get_char(word, char_index=0):
    """
    获取词语指定位置的字符（汉字）

    Args:
        word: 词语字符串
        char_index: 字符索引位置

    Returns:
        str: 指定位置的汉字，如果不是汉字返回None
    """
    if not word or len(word) <= char_index:
        return None

    target_char = word[char_index]

    # 检查是否为汉字
    if '\u4e00' <= target_char <= '\u9fff':
        return target_char

    return None


def analyze_special_words(classification):
    """
    分析特殊词汇的读音情况

    Args:
        classification: 词汇分类结果

    Returns:
        dict: 特殊词汇分析结果
    """
    special_analysis = {
        'single_chinese_analysis': [],
        'english_words_analysis': [],
        'mixed_words_analysis': [],
        'other_special_analysis': []
    }

    # 分析单字汉字词汇
    for word in classification['single_chinese']:
        pinyin = get_char_pinyin(word, 0)
        special_analysis['single_chinese_analysis'].append({
            'word': word,
            'pronunciation': word,  # 单字直接使用自身作为读音
            'pinyin': pinyin,
            'reason': '单字汉字，直接使用自身'
        })

    # 分析英文词汇
    for word in classification['english_words']:
        special_analysis['english_words_analysis'].append({
            'word': word,
            'pronunciation': word,  # 英文词汇使用自身
            'pinyin': None,
            'reason': '英文词汇，使用原词'
        })

    # 分析中英混合词汇
    for word in classification['mixed_words']:
        # 尝试提取汉字部分的拼音
        chinese_chars = [char for char in word if '\u4e00' <= char <= '\u9fff']
        if chinese_chars:
            first_chinese_pinyin = get_char_pinyin(chinese_chars[0], 0) if chinese_chars else None
            special_analysis['mixed_words_analysis'].append({
                'word': word,
                'pronunciation': word,  # 混合词汇使用完整词
                'pinyin': first_chinese_pinyin,
                'chinese_part': ''.join(chinese_chars),
                'reason': '中英混合词汇，使用完整词'
            })
        else:
            special_analysis['mixed_words_analysis'].append({
                'word': word,
                'pronunciation': word,
                'pinyin': None,
                'chinese_part': '',
                'reason': '中英混合词汇，使用完整词'
            })

    # 分析其他特殊词汇
    for word in classification['special_words']:
        special_analysis['other_special_analysis'].append({
            'word': word,
            'pronunciation': word,  # 特殊词汇使用自身
            'pinyin': None,
            'reason': '特殊词汇（数字/符号等），使用原词'
        })

    return special_analysis


# ==================== 第一类方案 ====================

def determine_word_pronunciation_method1(words):
    """
    第一类方案：根据首字读音的独特性确定读音表示
    只处理正常的多字汉字词汇
    """
    print("🔤 【第一类方案】正在获取所有词的首字读音...")

    # 先进行词汇分类
    classification = classify_words(words)
    normal_words = classification['normal_chinese']

    print(f"📊 词汇分类统计:")
    print(f"   正常汉字词汇: {len(normal_words)}")
    print(f"   单字汉字词汇: {len(classification['single_chinese'])}")
    print(f"   英文词汇: {len(classification['english_words'])}")
    print(f"   中英混合词汇: {len(classification['mixed_words'])}")
    print(f"   其他特殊词汇: {len(classification['special_words'])}")

    # 第一步：获取正常词汇的首字读音和首字
    word_info = []
    for word in normal_words:
        first_char = get_char(word, 0)
        first_char_pinyin = get_char_pinyin(word, 0)

        word_info.append({
            'word': word,
            'first_char': first_char,
            'first_char_pinyin': first_char_pinyin
        })

    # 过滤掉无法获取拼音的词
    valid_word_info = [info for info in word_info if info['first_char_pinyin'] is not None]

    print(f"✅ 成功获取 {len(valid_word_info)} 个正常词语的首字读音")

    # 第二步：统计首字读音的重复情况
    print("📊 正在统计首字读音重复情况...")

    pinyin_counter = Counter()
    for info in valid_word_info:
        pinyin_counter[info['first_char_pinyin']] += 1

    # 第三步：根据逻辑确定每个词的读音表示
    print("🎯 正在确定每个词的读音表示...")

    pronunciation_results = []

    for info in valid_word_info:
        word = info['word']
        first_char = info['first_char']
        first_char_pinyin = info['first_char_pinyin']

        # 检查首字读音是否独一无二
        if pinyin_counter[first_char_pinyin] == 1:
            # 独一无二，使用首字
            pronunciation = first_char
            reason = "首字读音独一无二"
            is_unique = True
        else:
            # 有重复，使用完整词语
            pronunciation = word
            reason = f"首字读音重复({pinyin_counter[first_char_pinyin]}次)"
            is_unique = False

        pronunciation_results.append({
            'word': word,
            'first_char': first_char,
            'first_char_pinyin': first_char_pinyin,
            'pronunciation': pronunciation,
            'is_unique': is_unique,
            'reason': reason,
            'repeat_count': pinyin_counter[first_char_pinyin]
        })

    # 统计信息
    total_words = len(pronunciation_results)
    unique_count = sum(1 for result in pronunciation_results if result['is_unique'])
    repeated_count = total_words - unique_count
    unique_pinyin_count = len(pinyin_counter)
    repeated_pinyin_count = sum(1 for count in pinyin_counter.values() if count > 1)

    stats = {
        'total_words': total_words,
        'unique_pronunciation_count': unique_count,
        'repeated_pronunciation_count': repeated_count,
        'unique_pinyin_count': unique_pinyin_count,
        'repeated_pinyin_count': repeated_pinyin_count,
        'pinyin_frequency': dict(pinyin_counter)
    }

    return pronunciation_results, stats, classification


# ==================== 第二类方案 ====================

def determine_word_pronunciation_method2(words):
    """
    第二类方案：基于首字和第二个字读音联合集合的独特性确定读音表示
    联合集合包含：正常词汇的首字读音 + 正常词汇的第二个字读音 + 单字词汇的读音
    """
    print("🔤 【第二类方案】正在获取所有词的首字和第二个字读音...")

    # 先进行词汇分类
    classification = classify_words(words)
    normal_words = classification['normal_chinese']
    single_words = classification['single_chinese']

    print(f"📊 词汇分类统计:")
    print(f"   正常汉字词汇: {len(normal_words)}")
    print(f"   单字汉字词汇: {len(single_words)}")

    # 第一步：获取所有词的首字和第二个字信息
    word_info = []
    all_first_pinyin = []
    all_second_pinyin = []
    all_single_pinyin = []  # 新增：单字词汇的读音

    # 处理正常词汇
    for word in normal_words:
        first_char = get_char(word, 0)
        first_char_pinyin = get_char_pinyin(word, 0)
        second_char = get_char(word, 1)
        second_char_pinyin = get_char_pinyin(word, 1)

        word_info.append({
            'word': word,
            'first_char': first_char,
            'first_char_pinyin': first_char_pinyin,
            'second_char': second_char,
            'second_char_pinyin': second_char_pinyin
        })

        # 收集所有拼音用于构建联合集合
        if first_char_pinyin:
            all_first_pinyin.append(first_char_pinyin)
        if second_char_pinyin:
            all_second_pinyin.append(second_char_pinyin)

    # 处理单字词汇，将其读音也加入联合集合
    for word in single_words:
        single_pinyin = get_char_pinyin(word, 0)
        if single_pinyin:
            all_single_pinyin.append(single_pinyin)

    # 过滤掉无法获取首字拼音的词
    valid_word_info = [info for info in word_info if info['first_char_pinyin'] is not None]

    print(f"✅ 成功获取 {len(valid_word_info)} 个正常词语的字符信息")
    print(f"✅ 成功获取 {len(all_single_pinyin)} 个单字词语的读音")

    # 第二步：构建联合集合并统计（包含单字读音）
    print("🔗 正在构建首字、第二个字和单字读音的联合集合...")

    # 构建联合集合（包含单字读音）
    union_pinyin_list = all_first_pinyin + all_second_pinyin + all_single_pinyin
    union_pinyin_counter = Counter(union_pinyin_list)

    # 分别统计各类读音
    first_pinyin_counter = Counter(all_first_pinyin)
    second_pinyin_counter = Counter(all_second_pinyin)
    single_pinyin_counter = Counter(all_single_pinyin)

    print(f"📊 联合集合统计:")
    print(f"   首字读音数: {len(all_first_pinyin)}")
    print(f"   第二个字读音数: {len(all_second_pinyin)}")
    print(f"   单字读音数: {len(all_single_pinyin)}")
    print(f"   联合集合总数: {len(union_pinyin_list)}")
    print(f"   联合集合唯一读音数: {len(union_pinyin_counter)}")

    # 第三步：根据逻辑确定每个词的读音表示
    print("🎯 正在根据联合集合确定每个词的读音表示...")

    pronunciation_results = []

    for info in valid_word_info:
        word = info['word']
        first_char = info['first_char']
        first_char_pinyin = info['first_char_pinyin']
        second_char = info['second_char']
        second_char_pinyin = info['second_char_pinyin']

        pronunciation = None
        reason = ""
        method_used = ""

        # 优先级1：检查首字在联合集合中是否独一无二
        if first_char_pinyin and union_pinyin_counter[first_char_pinyin] == 1:
            pronunciation = first_char
            reason = "首字读音在联合集合中独一无二"
            method_used = "首字"

        # 优先级2：检查第二个字在联合集合中是否独一无二
        elif second_char_pinyin and union_pinyin_counter[second_char_pinyin] == 1:
            pronunciation = second_char
            reason = "第二个字读音在联合集合中独一无二"
            method_used = "第二个字"

        # 优先级3：使用完整词汇
        else:
            pronunciation = word
            if first_char_pinyin and second_char_pinyin:
                reason = f"首字和第二个字读音都不独特(首字:{union_pinyin_counter[first_char_pinyin]}次, 第二个字:{union_pinyin_counter[second_char_pinyin]}次)"
            elif first_char_pinyin:
                reason = f"首字读音不独特({union_pinyin_counter[first_char_pinyin]}次), 第二个字无读音"
            else:
                reason = "首字读音不独特, 第二个字无读音"
            method_used = "完整词"

        pronunciation_results.append({
            'word': word,
            'first_char': first_char,
            'first_char_pinyin': first_char_pinyin,
            'second_char': second_char,
            'second_char_pinyin': second_char_pinyin,
            'pronunciation': pronunciation,
            'method_used': method_used,
            'reason': reason,
            'first_union_count': union_pinyin_counter.get(first_char_pinyin, 0),
            'second_union_count': union_pinyin_counter.get(second_char_pinyin, 0) if second_char_pinyin else 0
        })

    # 统计信息
    total_words = len(pronunciation_results)
    first_char_count = sum(1 for result in pronunciation_results if result['method_used'] == '首字')
    second_char_count = sum(1 for result in pronunciation_results if result['method_used'] == '第二个字')
    full_word_count = sum(1 for result in pronunciation_results if result['method_used'] == '完整词')

    stats = {
        'total_words': total_words,
        'first_char_count': first_char_count,
        'second_char_count': second_char_count,
        'full_word_count': full_word_count,
        'union_pinyin_count': len(union_pinyin_counter),
        'single_pinyin_count': len(all_single_pinyin),
        'first_pinyin_frequency': dict(first_pinyin_counter),
        'second_pinyin_frequency': dict(second_pinyin_counter),
        'single_pinyin_frequency': dict(single_pinyin_counter),
        'union_pinyin_frequency': dict(union_pinyin_counter)
    }

    return pronunciation_results, stats, classification


# ==================== 结果展示函数 ====================

def print_method1_results(pronunciation_results, stats):
    """
    打印第一类方案的分析结果
    """
    print("\n" + "=" * 80)
    print("【第一类方案】词语读音分析结果")
    print("=" * 80)

    # 基本统计
    print(f"\n📊 基本统计:")
    print(f"   总词数: {stats['total_words']}")
    print(
        f"   使用首字作为读音的词数: {stats['unique_pronunciation_count']} ({stats['unique_pronunciation_count'] / stats['total_words'] * 100:.1f}%)")
    print(
        f"   使用完整词作为读音的词数: {stats['repeated_pronunciation_count']} ({stats['repeated_pronunciation_count'] / stats['total_words'] * 100:.1f}%)")
    print(f"   唯一首字读音数: {stats['unique_pinyin_count']}")
    print(f"   重复首字读音种类数: {stats['repeated_pinyin_count']}")

    # 显示使用首字作为读音的词语示例
    unique_examples = [result for result in pronunciation_results if result['is_unique']]
    if unique_examples:
        print(f"\n✨ 使用首字作为读音的词语示例 (前15个):")
        for i, result in enumerate(unique_examples[:15]):
            print(f"   {i + 1:2d}. {result['word']:8s} → {result['pronunciation']} ({result['first_char_pinyin']})")

    # 显示使用完整词作为读音的词语示例
    repeated_examples = [result for result in pronunciation_results if not result['is_unique']]
    if repeated_examples:
        print(f"\n🔁 使用完整词作为读音的词语示例 (前15个):")
        for i, result in enumerate(repeated_examples[:15]):
            print(
                f"   {i + 1:2d}. {result['word']:8s} → {result['pronunciation']} ({result['first_char_pinyin']}, {result['repeat_count']}次重复)")


def print_method2_results(pronunciation_results, stats):
    """
    打印第二类方案的分析结果
    """
    print("\n" + "=" * 80)
    print("【第二类方案】词语读音分析结果")
    print("=" * 80)

    # 基本统计
    print(f"\n📊 基本统计:")
    print(f"   总词数: {stats['total_words']}")
    print(
        f"   使用首字作为读音的词数: {stats['first_char_count']} ({stats['first_char_count'] / stats['total_words'] * 100:.1f}%)")
    print(
        f"   使用第二个字作为读音的词数: {stats['second_char_count']} ({stats['second_char_count'] / stats['total_words'] * 100:.1f}%)")
    print(
        f"   使用完整词作为读音的词数: {stats['full_word_count']} ({stats['full_word_count'] / stats['total_words'] * 100:.1f}%)")
    print(f"   联合集合唯一读音数: {stats['union_pinyin_count']}")

    # 按方法分类显示示例
    method_examples = {
        '首字': [result for result in pronunciation_results if result['method_used'] == '首字'],
        '第二个字': [result for result in pronunciation_results if result['method_used'] == '第二个字'],
        '完整词': [result for result in pronunciation_results if result['method_used'] == '完整词']
    }

    for method, examples in method_examples.items():
        if examples:
            print(f"\n🎯 使用{method}作为读音的词语示例 (前10个):")
            for i, result in enumerate(examples[:10]):
                if method == '首字':
                    print(
                        f"   {i + 1:2d}. {result['word']:8s} → {result['pronunciation']} (首字:{result['first_char_pinyin']}, 联合集合中1次)")
                elif method == '第二个字':
                    print(
                        f"   {i + 1:2d}. {result['word']:8s} → {result['pronunciation']} (第二个字:{result['second_char_pinyin']}, 联合集合中1次)")
                else:
                    print(
                        f"   {i + 1:2d}. {result['word']:8s} → {result['pronunciation']} (首字:{result['first_union_count']}次, 第二个字:{result['second_union_count']}次)")


def print_special_words_analysis(special_analysis):
    """
    打印特殊词汇分析结果
    """
    print("\n" + "=" * 80)
    print("【特殊词汇分析】")
    print("=" * 80)

    # 单字汉字词汇
    if special_analysis['single_chinese_analysis']:
        print(f"\n📝 单字汉字词汇 ({len(special_analysis['single_chinese_analysis'])}个):")
        for i, item in enumerate(special_analysis['single_chinese_analysis'][:20]):
            print(f"   {i + 1:2d}. {item['word']} → {item['pronunciation']} ({item['pinyin']})")
        if len(special_analysis['single_chinese_analysis']) > 20:
            print(f"   ... 还有{len(special_analysis['single_chinese_analysis']) - 20}个")

    # 英文词汇
    if special_analysis['english_words_analysis']:
        print(f"\n🔤 英文词汇 ({len(special_analysis['english_words_analysis'])}个):")
        for i, item in enumerate(special_analysis['english_words_analysis'][:20]):
            print(f"   {i + 1:2d}. {item['word']} → {item['pronunciation']}")
        if len(special_analysis['english_words_analysis']) > 20:
            print(f"   ... 还有{len(special_analysis['english_words_analysis']) - 20}个")

    # 中英混合词汇
    if special_analysis['mixed_words_analysis']:
        print(f"\n🔀 中英混合词汇 ({len(special_analysis['mixed_words_analysis'])}个):")
        for i, item in enumerate(special_analysis['mixed_words_analysis'][:20]):
            chinese_part = f", 汉字部分:{item['chinese_part']}" if item['chinese_part'] else ""
            pinyin_part = f", 拼音:{item['pinyin']}" if item['pinyin'] else ""
            print(f"   {i + 1:2d}. {item['word']} → {item['pronunciation']}{chinese_part}{pinyin_part}")
        if len(special_analysis['mixed_words_analysis']) > 20:
            print(f"   ... 还有{len(special_analysis['mixed_words_analysis']) - 20}个")

    # 其他特殊词汇
    if special_analysis['other_special_analysis']:
        print(f"\n🔣 其他特殊词汇 ({len(special_analysis['other_special_analysis'])}个):")
        for i, item in enumerate(special_analysis['other_special_analysis'][:20]):
            print(f"   {i + 1:2d}. {item['word']} → {item['pronunciation']}")
        if len(special_analysis['other_special_analysis']) > 20:
            print(f"   ... 还有{len(special_analysis['other_special_analysis']) - 20}个")


def print_pronunciation_list(pronunciation_results, method_name):
    """
    打印读音列表
    """
    print(f"\n📝 【{method_name}】每个词的读音表示（汉字形式）:")
    print(f"{'序号':<4} {'原词':<10} {'读音':<10} {'方法':<8}")
    print("-" * 40)

    for i, result in enumerate(pronunciation_results[:30]):  # 只显示前30个
        method = result.get('method_used', '首字' if result.get('is_unique') else '完整词')
        print(f"{i + 1:<4} {result['word']:<10} {result['pronunciation']:<10} {method:<8}")

    if len(pronunciation_results) > 30:
        print(f"... 还有{len(pronunciation_results) - 30}个")

    # 提取纯读音列表
    pronunciation_list = [result['pronunciation'] for result in pronunciation_results]

    print(f"\n📋 【{method_name}】纯读音数组 (前30个):")
    print(pronunciation_list[:30])

    return pronunciation_list


def compare_methods(method1_results, method2_results):
    """
    对比两种方案的结果
    """
    print("\n" + "=" * 80)
    print("【方案对比分析】")
    print("=" * 80)

    method1_pronunciations = [result['pronunciation'] for result in method1_results]
    method2_pronunciations = [result['pronunciation'] for result in method2_results]

    # 统计差异
    different_count = 0
    same_count = 0
    differences = []

    for i, (m1_result, m2_result) in enumerate(zip(method1_results, method2_results)):
        if m1_result['pronunciation'] != m2_result['pronunciation']:
            different_count += 1
            differences.append({
                'index': i,
                'word': m1_result['word'],
                'method1_pronunciation': m1_result['pronunciation'],
                'method2_pronunciation': m2_result['pronunciation'],
                'method1_reason': m1_result['reason'],
                'method2_reason': m2_result['reason']
            })
        else:
            same_count += 1

    print(f"\n📊 对比统计:")
    print(f"   总词数: {len(method1_results)}")
    print(f"   读音相同的词数: {same_count} ({same_count / len(method1_results) * 100:.1f}%)")
    print(f"   读音不同的词数: {different_count} ({different_count / len(method1_results) * 100:.1f}%)")

    if differences:
        print(f"\n🔄 读音不同的词语示例 (前15个):")
        print(f"{'序号':<4} {'词语':<8} {'方案1读音':<10} {'方案2读音':<10}")
        print("-" * 50)
        for i, diff in enumerate(differences[:15]):
            print(
                f"{i + 1:<4} {diff['word']:<8} {diff['method1_pronunciation']:<10} {diff['method2_pronunciation']:<10}")

    return differences


def save_comprehensive_results(method1_results, method1_stats, method2_results, method2_stats,
                               differences, special_analysis, classification,
                               filename="comprehensive_pronunciation_results.txt"):
    """
    保存综合结果到文件，包含特殊词汇分析
    """
    try:
        with open(filename, 'w', encoding='utf-8') as f:
            f.write("词语读音分析综合结果\n")
            f.write("=" * 60 + "\n\n")

            # 词汇分类统计
            f.write("【词汇分类统计】\n")
            f.write(f"正常汉字词汇: {len(classification['normal_chinese'])}\n")
            f.write(f"单字汉字词汇: {len(classification['single_chinese'])}\n")
            f.write(f"英文词汇: {len(classification['english_words'])}\n")
            f.write(f"中英混合词汇: {len(classification['mixed_words'])}\n")
            f.write(f"其他特殊词汇: {len(classification['special_words'])}\n\n")

            # 第一类方案结果
            f.write("【第一类方案结果】(仅处理正常汉字词汇)\n")
            f.write(f"总词数: {method1_stats['total_words']}\n")
            f.write(
                f"使用首字: {method1_stats['unique_pronunciation_count']} ({method1_stats['unique_pronunciation_count'] / method1_stats['total_words'] * 100:.1f}%)\n")
            f.write(
                f"使用完整词: {method1_stats['repeated_pronunciation_count']} ({method1_stats['repeated_pronunciation_count'] / method1_stats['total_words'] * 100:.1f}%)\n\n")

            # 第二类方案结果
            f.write("【第二类方案结果】(仅处理正常汉字词汇)\n")
            f.write(f"总词数: {method2_stats['total_words']}\n")
            f.write(
                f"使用首字: {method2_stats['first_char_count']} ({method2_stats['first_char_count'] / method2_stats['total_words'] * 100:.1f}%)\n")
            f.write(
                f"使用第二个字: {method2_stats['second_char_count']} ({method2_stats['second_char_count'] / method2_stats['total_words'] * 100:.1f}%)\n")
            f.write(
                f"使用完整词: {method2_stats['full_word_count']} ({method2_stats['full_word_count'] / method2_stats['total_words'] * 100:.1f}%)\n\n")

            # 对比结果
            f.write("【方案对比】\n")
            f.write(f"读音相同的词数: {len(method1_results) - len(differences)}\n")
            f.write(f"读音不同的词数: {len(differences)}\n\n")

            # ==================== 特殊词汇分析 ====================
            f.write("【特殊词汇读音分析】\n")
            f.write("=" * 40 + "\n\n")

            # 单字汉字词汇
            if special_analysis['single_chinese_analysis']:
                f.write(f"单字汉字词汇 ({len(special_analysis['single_chinese_analysis'])}个):\n")
                f.write("序号\t词汇\t读音\t拼音\t说明\n")
                for i, item in enumerate(special_analysis['single_chinese_analysis']):
                    pinyin_str = item['pinyin'] if item['pinyin'] else "N/A"
                    f.write(f"{i + 1}\t{item['word']}\t{item['pronunciation']}\t{pinyin_str}\t{item['reason']}\n")
                f.write("\n")

            # 英文词汇
            if special_analysis['english_words_analysis']:
                f.write(f"英文词汇 ({len(special_analysis['english_words_analysis'])}个):\n")
                f.write("序号\t词汇\t读音\t说明\n")
                for i, item in enumerate(special_analysis['english_words_analysis']):
                    f.write(f"{i + 1}\t{item['word']}\t{item['pronunciation']}\t{item['reason']}\n")
                f.write("\n")

            # 中英混合词汇
            if special_analysis['mixed_words_analysis']:
                f.write(f"中英混合词汇 ({len(special_analysis['mixed_words_analysis'])}个):\n")
                f.write("序号\t词汇\t读音\t汉字部分\t汉字拼音\t说明\n")
                for i, item in enumerate(special_analysis['mixed_words_analysis']):
                    chinese_part = item['chinese_part'] if item['chinese_part'] else "无"
                    pinyin_str = item['pinyin'] if item['pinyin'] else "N/A"
                    f.write(
                        f"{i + 1}\t{item['word']}\t{item['pronunciation']}\t{chinese_part}\t{pinyin_str}\t{item['reason']}\n")
                f.write("\n")

            # 其他特殊词汇
            if special_analysis['other_special_analysis']:
                f.write(f"其他特殊词汇 ({len(special_analysis['other_special_analysis'])}个):\n")
                f.write("序号\t词汇\t读音\t说明\n")
                for i, item in enumerate(special_analysis['other_special_analysis']):
                    f.write(f"{i + 1}\t{item['word']}\t{item['pronunciation']}\t{item['reason']}\n")
                f.write("\n")

            # ==================== 正常词汇详细数据 ====================
            f.write("【正常汉字词汇详细数据】\n")
            f.write("=" * 40 + "\n")
            f.write("序号\t原词\t方案1读音\t方案2读音\t方案1方法\t方案2方法\n")
            for i, (m1_result, m2_result) in enumerate(zip(method1_results, method2_results)):
                m1_method = "首字" if m1_result.get('is_unique') else "完整词"
                m2_method = m2_result.get('method_used', '')
                f.write(
                    f"{i + 1}\t{m1_result['word']}\t{m1_result['pronunciation']}\t{m2_result['pronunciation']}\t{m1_method}\t{m2_method}\n")

            # ==================== 完整读音列表 ====================
            f.write("\n【完整读音列表】\n")
            f.write("=" * 40 + "\n")

            # 合并所有读音（正常词汇 + 特殊词汇）
            all_pronunciations_m1 = []
            all_pronunciations_m2 = []

            # 添加正常词汇的读音
            all_pronunciations_m1.extend([result['pronunciation'] for result in method1_results])
            all_pronunciations_m2.extend([result['pronunciation'] for result in method2_results])

            # 添加特殊词汇的读音
            for category in ['single_chinese_analysis', 'english_words_analysis', 'mixed_words_analysis',
                             'other_special_analysis']:
                for item in special_analysis[category]:
                    all_pronunciations_m1.append(item['pronunciation'])
                    all_pronunciations_m2.append(item['pronunciation'])

            f.write("方案1完整读音列表:\n")
            for i, pronunciation in enumerate(all_pronunciations_m1):
                f.write(f"{i + 1}. {pronunciation}\n")

            f.write("\n方案2完整读音列表:\n")
            for i, pronunciation in enumerate(all_pronunciations_m2):
                f.write(f"{i + 1}. {pronunciation}\n")

        print(f"\n💾 综合结果已保存到文件: {filename}")

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

    # 读取Excel数据（自动排除"-"）
    words = read_excel_range_by_columns(file_path, sheet_name, 'B2', 'Y25')

    if not words:
        print("❌ 没有读取到数据，请检查文件路径和工作表名称")
        return

    print(f"✅ 成功读取到 {len(words)} 个有效词语（已排除'-'）")

    # ==================== 运行第一类方案 ====================
    print("\n" + "🔵" * 20 + " 第一类方案 " + "🔵" * 20)
    method1_results, method1_stats, classification = determine_word_pronunciation_method1(words)

    if not method1_results:
        print("❌ 第一类方案没有找到有效的汉字词语")
        return

    print_method1_results(method1_results, method1_stats)
    method1_pronunciation_list = print_pronunciation_list(method1_results, "第一类方案")

    # ==================== 运行第二类方案 ====================
    print("\n" + "🟢" * 20 + " 第二类方案 " + "🟢" * 20)
    method2_results, method2_stats, _ = determine_word_pronunciation_method2(words)

    if not method2_results:
        print("❌ 第二类方案没有找到有效的汉字词语")
        return

    print_method2_results(method2_results, method2_stats)
    method2_pronunciation_list = print_pronunciation_list(method2_results, "第二类方案")

    # ==================== 特殊词汇分析 ====================
    print("\n" + "🟡" * 20 + " 特殊词汇分析 " + "🟡" * 20)
    special_analysis = analyze_special_words(classification)
    print_special_words_analysis(special_analysis)

    # ==================== 对比分析 ====================
    differences = compare_methods(method1_results, method2_results)

    # 保存综合结果（包含特殊词汇）
    save_comprehensive_results(method1_results, method1_stats, method2_results, method2_stats,
                               differences, special_analysis, classification)

    # 返回结果
    return {
        'words': words,
        'classification': classification,
        'method1_pronunciations': method1_pronunciation_list,
        'method2_pronunciations': method2_pronunciation_list,
        'method1_results': method1_results,
        'method2_results': method2_results,
        'method1_stats': method1_stats,
        'method2_stats': method2_stats,
        'special_analysis': special_analysis,
        'differences': differences
    }


if __name__ == "__main__":
    # 安装依赖提示
    print("📦 请确保已安装依赖: pip install pandas openpyxl pypinyin")
    print()

    # 运行脚本
    results = main()

    if results:
        print(f"\n🎉 综合分析完成！")
        print(f"   - 总词语数: {len(results['words'])}")
        print(f"   - 正常汉字词语: {len(results['classification']['normal_chinese'])}")
        print(f"   - 单字汉字词语: {len(results['classification']['single_chinese'])}")
        print(f"   - 英文词语: {len(results['classification']['english_words'])}")
        print(f"   - 中英混合词语: {len(results['classification']['mixed_words'])}")
        print(f"   - 其他特殊词语: {len(results['classification']['special_words'])}")
        print(f"   - 两方案差异词语数: {len(results['differences'])}")

        # 计算完整读音数组长度（包含特殊词汇）
        total_special = sum(len(results['special_analysis'][key]) for key in results['special_analysis'])
        total_pronunciations = len(results['method1_pronunciations']) + total_special

        print(f"\n📊 完整读音数组:")
        print(f"   - 第一类方案完整读音数组长度: {total_pronunciations}")
        print(f"   - 第二类方案完整读音数组长度: {total_pronunciations}")