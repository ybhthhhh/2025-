import pandas as pd
import numpy as np
from datetime import datetime
import random

# 设置随机种子以确保结果可重现
random.seed(42)
np.random.seed(42)

def preprocess_data(input_file):
    """
    第一步：预处理数据，删除重复的人，只保留最后一次填写
    """
    # 读取Excel文件
    df = pd.read_excel(input_file, sheet_name='Sheet1')
    
    # 按提交时间排序，确保最新的记录在后面
    df['提交答卷时间'] = pd.to_datetime(df['提交答卷时间'])
    df = df.sort_values('提交答卷时间')
    
    # 删除完全重复的行
    df = df.drop_duplicates()
    
    # 按姓名和手机号分组，保留最后一条记录
    df_preprocessed = df.drop_duplicates(
        subset=['1、您的姓名是？', '3、您的手机号码是？'], 
        keep='last'
    )
    
    # 保存预处理后的数据
    df_preprocessed.to_excel('excel1_预处理后的数据.xlsx', index=False)
    print("第一步完成：预处理数据已保存为 excel1_预处理后的数据.xlsx")
    
    return df_preprocessed

def split_data(df):
    """
    第二步：将数据分成匹配舞伴和自带舞伴两组
    """
    # 匹配舞伴的同学
    df_match = df[df['10、您选择自带舞伴还是匹配舞伴？（如自带舞伴两人都要填写问卷且信息务必对应一致）'] == '匹配舞伴'].copy()
    df_match.to_excel('excel2_匹配舞伴的同学.xlsx', index=False)
    
    # 自带舞伴的同学
    df_bring = df[df['10、您选择自带舞伴还是匹配舞伴？（如自带舞伴两人都要填写问卷且信息务必对应一致）'] == '自带舞伴'].copy()
    df_bring.to_excel('excel3_自带舞伴的同学.xlsx', index=False)
    
    print("第二步完成：数据已分割为匹配舞伴和自带舞伴两组")
    print(f"匹配舞伴人数: {len(df_match)}")
    print(f"自带舞伴人数: {len(df_bring)}")
    
    return df_match, df_bring

def is_gender_match(gender_pref, actual_gender):
    """
    检查性别偏好是否匹配
    """
    if gender_pref in ['(跳过)', '随便', '均可', '任意', '都行']:
        return True
    return gender_pref == actual_gender

def calculate_compatibility_score(person1, person2):
    """
    计算两个人之间的匹配度分数
    """
    score = 0
    
    # 检查性别要求是否符合 (必须满足，否则返回0分)
    gender_pref1 = person1['15、您希望舞伴的性别是？']
    gender_pref2 = person2['15、您希望舞伴的性别是？']
    gender1 = person1['2、您的性别是？']
    gender2 = person2['2、您的性别是？']
    
    # 检查双方的性别偏好是否匹配
    if not is_gender_match(gender_pref1, gender2) or not is_gender_match(gender_pref2, gender1):
        return 0
    
    # 检查是否同为本科生 (5分)
    if person1['5、您是否为本科生？'] == person2['5、您是否为本科生？']:
        score += 5
    
    # 检查是否同为研究生 (如果不是本科生，假设是研究生) (5分)
    if person1['5、您是否为本科生？'] != '是' and person2['5、您是否为本科生？'] != '是':
        score += 5
    
    # 检查是否同为新生 (5分)
    if person1['6、您是否为25级新生？'] == person2['6、您是否为25级新生？']:
        score += 5
    
    # 检查外向内向匹配 (10分)
    if person1['17、您希望您的舞伴是内向还是外向？'] == '随便' or person2['17、您希望您的舞伴是内向还是外向？'] == '随便':
        score += 5  # 如果有人不介意，给中等分数
    elif person1['16、您是外向还是内向？'] == person2['17、您希望您的舞伴是内向还是外向？'] and \
         person2['16、您是外向还是内向？'] == person1['17、您希望您的舞伴是内向还是外向？']:
        score += 10  # 完美匹配
    
    # 相同院系额外加分 (3分)
    if person1['8、您所在的院系是？'] == person2['8、您所在的院系是？']:
        score += 3
    
    return score

def match_partners(df_match):
    """
    第三步：使用改进的贪心算法匹配舞伴，支持各种性别组合
    """
    if len(df_match) < 2:
        return pd.DataFrame()
    
    # 创建所有人列表并重置索引
    people = df_match.reset_index(drop=True).to_dict('records')
    
    # 创建匹配结果列表
    matches = []
    matched_indices = set()
    
    # 正确初始化兼容性矩阵
    compatibility_matrix = {}
    for i in range(len(people)):
        compatibility_matrix[i] = {}
    
    # 为每个人计算与其他所有人的匹配度
    for i in range(len(people)):
        for j in range(i + 1, len(people)):
            score = calculate_compatibility_score(people[i], people[j])
            if score > 0:  # 只考虑有效匹配
                compatibility_matrix[i][j] = score
                compatibility_matrix[j][i] = score
    
    # 按照匹配度分数排序所有可能的配对
    all_possible_pairs = []
    for i in range(len(people)):
        for j in range(i + 1, len(people)):
            if j in compatibility_matrix.get(i, {}):
                all_possible_pairs.append((i, j, compatibility_matrix[i][j]))
    
    # 按匹配度分数降序排序
    all_possible_pairs.sort(key=lambda x: x[2], reverse=True)
    
    # 贪心算法：优先匹配分数高的配对
    for i, j, score in all_possible_pairs:
        if i not in matched_indices and j not in matched_indices:
            matches.append((people[i], people[j], score))
            matched_indices.add(i)
            matched_indices.add(j)
    
    # 创建匹配结果DataFrame
    matched_pairs = []
    for person1, person2, score in matches:
        # 确定谁作为person1（按字母顺序排序，便于阅读）
        if person1['1、您的姓名是？'] < person2['1、您的姓名是？']:
            first_person, second_person = person1, person2
        else:
            first_person, second_person = person2, person1
        
        pair_data = {
            '姓名_1': first_person['1、您的姓名是？'],
            '性别_1': first_person['2、您的性别是？'],
            '手机号_1': first_person['3、您的手机号码是？'],
            '邮箱_1': first_person['4、您的邮箱是？（若为北大学生，建议填写北大邮箱）'],
            '院系_1': first_person['8、您所在的院系是？'],
            '舞会志愿顺序_1': first_person['9、您倾向于参加哪天的舞会？'],
            '外向内向_1': first_person['16、您是外向还是内向？'],
            '希望舞伴性别_1': first_person['15、您希望舞伴的性别是？'],
            
            '姓名_2': second_person['1、您的姓名是？'],
            '性别_2': second_person['2、您的性别是？'],
            '手机号_2': second_person['3、您的手机号码是？'],
            '邮箱_2': second_person['4、您的邮箱是？（若为北大学生，建议填写北大邮箱）'],
            '院系_2': second_person['8、您所在的院系是？'],
            '舞会志愿顺序_2': second_person['9、您倾向于参加哪天的舞会？'],
            '外向内向_2': second_person['16、您是外向还是内向？'],
            '希望舞伴性别_2': second_person['15、您希望舞伴的性别是？'],
            
            '匹配度分数': score,
            '匹配类型': '算法匹配',
            '配对组合': f"{first_person['2、您的性别是？']}{second_person['2、您的性别是？']}"
        }
        matched_pairs.append(pair_data)
    
    # 添加未匹配的人员信息
    unmatched_people = []
    for i, person in enumerate(people):
        if i not in matched_indices:
            unmatched_people.append({
                '姓名': person['1、您的姓名是？'],
                '性别': person['2、您的性别是？'],
                '手机号': person['3、您的手机号码是？'],
                '希望舞伴性别': person['15、您希望舞伴的性别是？'],
                '外向内向': person['16、您是外向还是内向？']
            })
    
    if unmatched_people:
        print(f"有 {len(unmatched_people)} 人未能匹配到舞伴")
    
    df_matched = pd.DataFrame(matched_pairs)
    return df_matched

def create_pair_data(person1, person2, match_type):
    """创建配对数据"""
    if person1['1、您的姓名是？'] < person2['1、您的姓名是？']:
        first, second = person1, person2
    else:
        first, second = person2, person1
    
    return {
        '姓名_1': first['1、您的姓名是？'],
        '性别_1': first['2、您的性别是？'],
        '手机号_1': first['3、您的手机号码是？'],
        '邮箱_1': first['4、您的邮箱是？（若为北大学生，建议填写北大邮箱）'],
        '院系_1': first['8、您所在的院系是？'],
        '舞会志愿顺序_1': first['9、您倾向于参加哪天的舞会？'],
        '外向内向_1': first['16、您是外向还是内向？'],
        '希望舞伴性别_1': first['15、您希望舞伴的性别是？'],
        
        '姓名_2': second['1、您的姓名是？'],
        '性别_2': second['2、您的性别是？'],
        '手机号_2': second['3、您的手机号码是？'],
        '邮箱_2': second['4、您的邮箱是？（若为北大学生，建议填写北大邮箱）'],
        '院系_2': second['8、您所在的院系是？'],
        '舞会志愿顺序_2': second['9、您倾向于参加哪天的舞会？'],
        '外向内向_2': second['16、您是外向还是内向？'],
        '希望舞伴性别_2': second['15、您希望舞伴的性别是？'],
        
        '匹配度分数': 'N/A',
        '匹配类型': match_type,
        '配对组合': f"{first['2、您的性别是？']}{second['2、您的性别是？']}"
    }

def verify_bring_partners(df_bring):
    """
    验证自带舞伴的配对
    """
    if df_bring.empty:
        return pd.DataFrame()
    
    # 数据清洗
    df_bring_clean = df_bring.copy()
    
    # 清洗电话号码和姓名
    for col in ['3、您的手机号码是？', '12、您舞伴的电话号码是？（为对方最后填写的电话号码）']:
        df_bring_clean[col] = df_bring_clean[col].astype(str).str.strip()
    
    for col in ['1、您的姓名是？', '11、您舞伴的姓名是？']:
        df_bring_clean[col] = df_bring_clean[col].astype(str).str.strip()
    
    verified_pairs = []
    matched_indices = set()
    
    # 创建索引字典
    person_dict = {}
    for idx, row in df_bring_clean.iterrows():
        key = (row['1、您的姓名是？'], row['3、您的手机号码是？'])
        person_dict[key] = (idx, row)
    
    # 尝试匹配
    for (name, phone), (idx, person) in person_dict.items():
        if idx in matched_indices:
            continue
            
        partner_name = person['11、您舞伴的姓名是？']
        partner_phone = person['12、您舞伴的电话号码是？（为对方最后填写的电话号码）']
        
        # 查找匹配的舞伴
        partner_key = (partner_name, partner_phone)
        if partner_key in person_dict:
            partner_idx, partner_person = person_dict[partner_key]
            
            # 检查舞伴是否也选择了此人
            partner_partner_name = partner_person['11、您舞伴的姓名是？']
            partner_partner_phone = partner_person['12、您舞伴的电话号码是？（为对方最后填写的电话号码）']
            
            if (partner_partner_name == name and 
                partner_partner_phone == phone):
                
                # 创建配对数据
                pair_data = create_pair_data(person, partner_person, "自带舞伴")
                verified_pairs.append(pair_data)
                
                matched_indices.add(idx)
                matched_indices.add(partner_idx)
    
    return pd.DataFrame(verified_pairs)

def assign_dance_sessions(df_matched):
    """
    第四步：分配舞会场次
    """
    if df_matched.empty:
        return pd.DataFrame()
    
    # 初始化三个场次的计数器
    session_counts = {'周五': 0, '周六': 0, '周日': 0}
    max_pairs_per_session = 110
    
    # 待定列表
    pending_pairs = []
    assigned_pairs = []
    
    # 首先处理所有配对
    for _, pair in df_matched.iterrows():
        # 获取双方的志愿顺序
        pref1 = pair['舞会志愿顺序_1']
        pref2 = pair['舞会志愿顺序_2']
        
        # 解析志愿顺序，提取第一志愿
        def get_first_preference(pref_str):
            if not isinstance(pref_str, str):
                return None
            
            if '10月10日（周五）' in pref_str:
                if pref_str.startswith('10月10日（周五）'):
                    return '周五'
            if '10月11日（周六）' in pref_str:
                if pref_str.startswith('10月11日（周六）'):
                    return '周六'
            if '10月12日（周日）' in pref_str:
                if pref_str.startswith('10月12日（周日）'):
                    return '周日'
            
            # 如果无法从开头判断，尝试查找优先顺序
            if '→' in pref_str:
                choices = pref_str.split('→')
                for choice in choices:
                    choice = choice.strip()
                    if '周五' in choice:
                        return '周五'
                    elif '周六' in choice:
                        return '周六'
                    elif '周日' in choice:
                        return '周日'
            
            return None
        
        first_pref1 = get_first_preference(pref1)
        first_pref2 = get_first_preference(pref2)
        
        # 决定分配场次
        assigned_session = None
        
        # 收集双方的所有偏好
        preferences = set()
        if first_pref1:
            preferences.add(first_pref1)
        if first_pref2:
            preferences.add(first_pref2)
        
        # 按优先级尝试分配场次
        for session in ['周五', '周六', '周日']:
            if session in preferences and session_counts[session] < max_pairs_per_session:
                assigned_session = session
                session_counts[session] += 1
                break
        
        # 如果没有匹配的偏好或者偏好场次已满，放入待定
        if assigned_session is None:
            pending_pairs.append((pair, list(preferences)))
            continue
        
        # 添加分配结果
        pair_with_session = pair.to_dict()
        pair_with_session['分配场次'] = assigned_session
        assigned_pairs.append(pair_with_session)
    
    # 处理待定配对
    for pair, preferences in pending_pairs:
        assigned_session = None
        
        # 首先尝试首选场次
        for session in preferences:
            if session_counts[session] < max_pairs_per_session:
                assigned_session = session
                session_counts[session] += 1
                break
        
        # 如果没有首选场次可用，随机分配还有空位的场次
        if assigned_session is None:
            available_sessions = [s for s in ['周五', '周六', '周日'] if session_counts[s] < max_pairs_per_session]
            if available_sessions:
                assigned_session = random.choice(available_sessions)
                session_counts[assigned_session] += 1
            else:
                # 所有场次都满了，分配到周五
                assigned_session = '周五'
                session_counts['周五'] += 1
        
        pair_with_session = pair.to_dict()
        pair_with_session['分配场次'] = assigned_session
        assigned_pairs.append(pair_with_session)
    
    # 创建分配结果DataFrame
    df_assigned = pd.DataFrame(assigned_pairs)
    
    # 按场次分割数据
    friday_pairs = df_assigned[df_assigned['分配场次'] == '周五']
    saturday_pairs = df_assigned[df_assigned['分配场次'] == '周六']
    sunday_pairs = df_assigned[df_assigned['分配场次'] == '周日']
    
    # 保存到Excel的不同sheet中
    with pd.ExcelWriter('excel5_舞会场次分配.xlsx') as writer:
        friday_pairs.to_excel(writer, sheet_name='周五场次', index=False)
        saturday_pairs.to_excel(writer, sheet_name='周六场次', index=False)
        sunday_pairs.to_excel(writer, sheet_name='周日场次', index=False)
    
    print("第四步完成：舞会场次分配已保存为 excel5_舞会场次分配.xlsx")
    print(f"周五场次: {len(friday_pairs)} 对")
    print(f"周六场次: {len(saturday_pairs)} 对")
    print(f"周日场次: {len(sunday_pairs)} 对")
    
    return df_assigned

def main():
    """
    主函数：执行所有步骤
    """
    input_file = "星河之约”新生舞会报名问卷.xlsx"
    
    try:
        # 第一步：预处理数据
        print("正在执行第一步：数据预处理...")
        df_preprocessed = preprocess_data(input_file)
        
        # 第二步：分割数据
        print("正在执行第二步：数据分割...")
        df_match, df_bring = split_data(df_preprocessed)
        
        # 第三步：匹配舞伴
        print("正在执行第三步：匹配舞伴...")
        
        # 匹配算法匹配的舞伴
        if not df_match.empty:
            df_matched_algorithm = match_partners(df_match)
        else:
            df_matched_algorithm = pd.DataFrame()
            print("没有需要算法匹配的舞伴")
        
        # 验证自带舞伴
        if not df_bring.empty:
            df_matched_bring = verify_bring_partners(df_bring)
        else:
            df_matched_bring = pd.DataFrame()
            print("没有自带舞伴需要验证")
        
        # 合并所有匹配结果
        if not df_matched_algorithm.empty and not df_matched_bring.empty:
            df_all_matched = pd.concat([df_matched_algorithm, df_matched_bring], ignore_index=True)
        elif not df_matched_algorithm.empty:
            df_all_matched = df_matched_algorithm
        elif not df_matched_bring.empty:
            df_all_matched = df_matched_bring
        else:
            df_all_matched = pd.DataFrame()
            print("没有匹配到任何舞伴对")
        
        if not df_all_matched.empty:
            df_all_matched.to_excel('excel4_所有匹配结果.xlsx', index=False)
            print("第三步完成：所有匹配结果已保存为 excel4_所有匹配结果.xlsx")
            print(f"总共匹配了 {len(df_all_matched)} 对舞伴")
            
            # 第四步：分配舞会场次
            print("正在执行第四步：分配舞会场次...")
            df_assigned = assign_dance_sessions(df_all_matched)
            
            print("所有步骤完成！")
            print(f"生成的文件: excel1_预处理后的数据.xlsx, excel2_匹配舞伴的同学.xlsx, excel3_自带舞伴的同学.xlsx, excel4_所有匹配结果.xlsx, excel5_舞会场次分配.xlsx")
        else:
            print("没有匹配到任何舞伴对，无法进行场次分配")
            
    except Exception as e:
        print(f"程序执行出错: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()