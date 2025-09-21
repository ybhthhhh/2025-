import pandas as pd
import numpy as np
from datetime import datetime
import random
import traceback

random.seed(42)
np.random.seed(42)

# ---------- 工具 ----------
def is_gender_match(gender_pref, actual_gender):
    if gender_pref in {'(跳过)', '随便', '均可', '任意', '都行'}:
        return True
    return gender_pref == actual_gender


def calculate_compatibility_score(person1, person2):
    score = 0
    g1, g2 = person1['2、您的性别是？'], person2['2、您的性别是？']
    gp1, gp2 = person1['15、您希望舞伴的性别是？'], person2['15、您希望舞伴的性别是？']
    if not (is_gender_match(gp1, g2) and is_gender_match(gp2, g1)):
        return 0

    if person1['5、您是否为本科生？'] == person2['5、您是否为本科生？']:
        score += 5
    if person1['6、您是否为25级新生？'] == '是' and person2['6、您是否为25级新生？'] == '是':
        score += 5

    ext1, ext2 = person1['16、您是外向还是内向？'], person2['16、您是外向还是内向？']
    want1, want2 = person1['17、您希望您的舞伴是内向还是外向？'], person2['17、您希望您的舞伴是内向还是外向？']
    if want1 == '随便' or want2 == '随便':
        score += 5
    elif ext1 == want2 and ext2 == want1:
        score += 10

    return score


def is_pku_student(person_data):
    university = person_data.get('大学', '') or person_data.get('7、您的大学是？', '')
    return '北京大学' in str(university)


# ---------- 核心步骤 ----------
def preprocess_data(input_file):
    df = pd.read_excel(input_file, sheet_name='Sheet1')
    df['提交答卷时间'] = pd.to_datetime(df['提交答卷时间'])
    df = df.sort_values('提交答卷时间').drop_duplicates()
    df = df.drop_duplicates(subset=['1、您的姓名是？', '3、您的手机号码是？'], keep='last')
    
    # 按照序号列排序
    if '序号' in df.columns:
        df = df.sort_values('序号')
        print("已按照'序号'列进行排序")
    else:
        print("警告：未找到'序号'列，保持原顺序")
    
    df.to_excel('excel1_预处理后的数据.xlsx', index=False)
    print("第一步完成：预处理数据已保存为 excel1_预处理后的数据.xlsx")
    return df


def split_data(df):
    df_match = df[df['10、您选择自带舞伴还是匹配舞伴？（如自带舞伴两人都要填写问卷且信息务必对应一致）'] == '匹配舞伴'].copy()
    df_bring = df[df['10、您选择自带舞伴还是匹配舞伴？（如自带舞伴两人都要填写问卷且信息务必对应一致）'] == '自带舞伴'].copy()
    df_match.to_excel('excel2_匹配舞伴的同学.xlsx', index=False)
    df_bring.to_excel('excel3_自带舞伴的同学.xlsx', index=False)
    print("第二步完成：数据已分割")
    return df_match, df_bring


def match_partners(df_match):
    if len(df_match) < 2:
        return pd.DataFrame()
    people = df_match.reset_index(drop=True).to_dict('records')
    matches, matched_indices = [], set()
    compat = {}
    for i in range(len(people)):
        compat[i] = {}
    for i in range(len(people)):
        for j in range(i + 1, len(people)):
            s = calculate_compatibility_score(people[i], people[j])
            if s > 0:
                compat[i][j] = s
                compat[j][i] = s
    pairs = [(i, j, compat[i][j]) for i in compat for j in compat[i] if i < j]
    pairs.sort(key=lambda x: x[2], reverse=True)
    for i, j, s in pairs:
        if i not in matched_indices and j not in matched_indices:
            matches.append((people[i], people[j], s))
            matched_indices.update([i, j])
    matched_rows = []
    for p1, p2, s in matches:
        if p1['1、您的姓名是？'] > p2['1、您的姓名是？']:
            p1, p2 = p2, p1
        matched_rows.append({
            '姓名_1': p1['1、您的姓名是？'], '性别_1': p1['2、您的性别是？'],
            '手机号_1': p1['3、您的手机号码是？'], '邮箱_1': p1['4、您的邮箱是？（若为北大学生，建议填写北大邮箱）'],
            '院系_1': p1['8、您所在的院系是？'], '舞会志愿顺序_1': p1['9、您倾向于参加哪天的舞会？'],
            '外向内向_1': p1['16、您是外向还是内向？'], '希望舞伴性别_1': p1['15、您希望舞伴的性别是？'],
            '大学_1': p1['7、您的大学是？'], '6、您是否为25级新生？_1': p1['6、您是否为25级新生？'],
            '姓名_2': p2['1、您的姓名是？'], '性别_2': p2['2、您的性别是？'],
            '手机号_2': p2['3、您的手机号码是？'], '邮箱_2': p2['4、您的邮箱是？（若为北大学生，建议填写北大邮箱）'],
            '院系_2': p2['8、您所在的院系是？'], '舞会志愿顺序_2': p2['9、您倾向于参加哪天的舞会？'],
            '外向内向_2': p2['16、您是外向还是内向？'], '希望舞伴性别_2': p2['15、您希望舞伴的性别是？'],
            '大学_2': p2['7、您的大学是？'], '6、您是否为25级新生？_2': p2['6、您是否为25级新生？'],
            '匹配度分数': s, '匹配类型': '算法匹配', '配对组合': f"{p1['2、您的性别是？']}{p2['2、您的性别是？']}"
        })
    return pd.DataFrame(matched_rows)


def verify_bring_partners(df_bring):
    if df_bring.empty:
        return pd.DataFrame()
    dfb = df_bring.copy()
    for c in ['3、您的手机号码是？', '12、您舞伴的电话号码是？（为对方最后填写的电话号码）']:
        dfb[c] = dfb[c].astype(str).str.strip()
    for c in ['1、您的姓名是？', '11、您舞伴的姓名是？']:
        dfb[c] = dfb[c].astype(str).str.strip()
    verified, used = [], set()
    key_to_row = {(r['1、您的姓名是？'], r['3、您的手机号码是？']): (i, r) for i, r in dfb.iterrows()}
    for (name, phone), (idx, row) in key_to_row.items():
        if idx in used:
            continue
        pn, pp = row['11、您舞伴的姓名是？'], row['12、您舞伴的电话号码是？（为对方最后填写的电话号码）']
        if (pn, pp) in key_to_row:
            j, jr = key_to_row[(pn, pp)]
            if jr['11、您舞伴的姓名是？'] == name and jr['12、您舞伴的电话号码是？（为对方最后填写的电话号码）'] == phone:
                p1, p2 = row, jr
                if p1['1、您的姓名是？'] > p2['1、您的姓名是？']:
                    p1, p2 = p2, p1
                verified.append({
                    '姓名_1': p1['1、您的姓名是？'], '性别_1': p1['2、您的性别是？'],
                    '手机号_1': p1['3、您的手机号码是？'], '邮箱_1': p1['4、您的邮箱是？（若为北大学生，建议填写北大邮箱）'],
                    '院系_1': p1['8、您所在的院系是？'], '舞会志愿顺序_1': p1['9、您倾向于参加哪天的舞会？'],
                    '外向内向_1': p1['16、您是外向还是内向？'], '希望舞伴性别_1': p1['15、您希望舞伴的性别是？'],
                    '大学_1': p1['7、您的大学是？'], '6、您是否为25级新生？_1': p1['6、您是否为25级新生？'],
                    '姓名_2': p2['1、您的姓名是？'], '性别_2': p2['2、您的性别是？'],
                    '手机号_2': p2['3、您的手机号码是？'], '邮箱_2': p2['4、您的邮箱是？（若为北大学生，建议填写北大邮箱）'],
                    '院系_2': p2['8、您所在的院系是？'], '舞会志愿顺序_2': p2['9、您倾向于参加哪天的舞会？'],
                    '外向内向_2': p2['16、您是外向还是内向？'], '希望舞伴性别_2': p2['15、您希望舞伴的性别是？'],
                    '大学_2': p2['7、您的大学是？'], '6、您是否为25级新生？_2': p2['6、您是否为25级新生？'],
                    '匹配度分数': 'N/A', '匹配类型': '自带舞伴', '配对组合': f"{p1['2、您的性别是？']}{p2['2、您的性别是？']}"
                })
                used.update([idx, j])
    return pd.DataFrame(verified)


# ---------- 场次分配 ----------
def assign_dance_sessions(df_all_matched, _):
    if df_all_matched.empty:
        return pd.DataFrame()

    session_counts = {'周五': 0, '周六': 0, '周日': 0}
    max_pairs = 115

    bring = df_all_matched[df_all_matched['匹配类型'] == '自带舞伴'].copy()
    algo  = df_all_matched[df_all_matched['匹配类型'] == '算法匹配'].copy()

    # 1. 筛北大自带 → 再分新生/非新生
    pku_bring = []
    for _, r in bring.iterrows():
        if is_pku_student({'大学': r.get('大学_1', '')}) or is_pku_student({'大学': r.get('大学_2', '')}):
            f1 = r.get('6、您是否为25级新生？_1', '') == '是'
            f2 = r.get('6、您是否为25级新生？_2', '') == '是'
            r['优先级'] = '新生自带' if (f1 or f2) else '非新生自带'
            pku_bring.append(r)

    # 2. 非新生自带 → 评分16 加入算法池
    non_fresh = [r for r in pku_bring if r['优先级'] == '非新生自带']
    for r in non_fresh:
        r['匹配度分数'] = 16
        r['匹配类型']   = '自带舞伴（非新生）'
    algo = pd.concat([algo, pd.DataFrame(non_fresh)], ignore_index=True)

    # 真正优先的只剩“新生自带”
    pku_bring = [r for r in pku_bring if r['优先级'] == '新生自带']
    print(f"北大+新生 自带: {len(pku_bring)} 对 | 北大+非新生 自带: {len(non_fresh)} 对 → 进算法池")

    assigned = []

    # 3. 优先分配新生自带
    for r in pku_bring:
        sess = assign_bring_pair_session(r, session_counts, max_pairs)
        if not sess:
            sess = assign_pending_bring_pair(r, session_counts, max_pairs)
        rd = r.to_dict()
        rd['分配场次'] = sess
        assigned.append(rd)

    # 4. 分配算法池（含非新生自带）
    for _, r in algo.iterrows():
        sess = assign_session_to_pair(r, session_counts, max_pairs)
        if not sess:
            sess = assign_pending_pair(r, session_counts, max_pairs)
        rd = r.to_dict()
        rd['分配场次'] = sess
        assigned.append(rd)

    df_assigned = pd.DataFrame(assigned)

    # 5. 写 Excel
    fri = df_assigned[df_assigned['分配场次'] == '周五']
    with pd.ExcelWriter('excel5_舞会场次分配.xlsx', mode='w', engine='openpyxl') as w:
        fri.to_excel(w, sheet_name='周五场次', index=False)
    for d in ['周六', '周日']:
        tmp = df_assigned[df_assigned['分配场次'] == d]
        with pd.ExcelWriter('excel5_舞会场次分配.xlsx', mode='a', engine='openpyxl', if_sheet_exists='replace') as w:
            tmp.to_excel(w, sheet_name=f'{d}场次', index=False)

    for d in ['周五', '周六', '周日']:
        print(f"{d}场次: {len(df_assigned[df_assigned['分配场次']==d])} 对 (容量: {max_pairs})")
    print("第四步完成：excel5_舞会场次分配.xlsx 已保存")
    return df_assigned


# ---------- 下面 4 个辅助函数 ----------
def assign_bring_pair_session(pair, cnt, max_p):
    pref1, pref2 = pair['舞会志愿顺序_1'], pair['舞会志愿顺序_2']
    for day in ['周五', '周六', '周日']:
        if has_session(pref1, day) and has_session(pref2, day) and cnt[day] < max_p:
            cnt[day] += 1
            return day
    return None


def assign_pending_bring_pair(pair, cnt, max_p):
    av = [d for d in ['周五', '周六', '周日'] if cnt[d] < max_p]
    if av:
        d = random.choice(av)
        cnt[d] += 1
        return d
    cnt['周五'] += 1
    return '周五'


def assign_session_to_pair(pair, cnt, max_p):
    p1, p2 = pair['舞会志愿顺序_1'], pair['舞会志愿顺序_2']
    prefs = get_preferences_from_strings(p1, p2)
    for day in prefs:
        if cnt[day] < max_p:
            cnt[day] += 1
            return day
    return None


def assign_pending_pair(pair, cnt, max_p):
    p1, p2 = pair['舞会志愿顺序_1'], pair['舞会志愿顺序_2']
    prefs = get_preferences_from_strings(p1, p2)
    for day in prefs:
        if cnt[day] < max_p:
            cnt[day] += 1
            return day
    av = [d for d in ['周五', '周六', '周日'] if cnt[d] < max_p]
    if av:
        d = random.choice(av)
        cnt[d] += 1
        return d
    cnt['周五'] += 1
    return '周五'


def has_session(s, day):
    if not isinstance(s, str):
        return False
    if day == '周五':
        return '10月10日（周五）' in s or '周五' in s
    if day == '周六':
        return '10月11日（周六）' in s or '周六' in s
    if day == '周日':
        return '10月12日（周日）' in s or '周日' in s
    return False


def get_preferences_from_strings(p1, p2):
    def first(x):
        if not isinstance(x, str):
            return None
        if '→' in x:
            x = x.split('→')[0].strip()
        if '周五' in x or '10月10日' in x:
            return '周五'
        if '周六' in x or '10月11日' in x:
            return '周六'
        if '周日' in x or '10月12日' in x:
            return '周日'
        return None
    f1, f2 = first(p1), first(p2)
    prefs = list(dict.fromkeys([f for f in [f1, f2] if f]))
    if '周五' in prefs:
        return ['周五']
    return prefs if prefs else ['周五', '周六', '周日']


# ---------- 主流程 ----------
def main():
    try:
        print("=== 第一步：预处理 ===")
        df_pre = preprocess_data("星河之约新生舞会报名问卷.xlsx")
        print("=== 第二步：分割 ===")
        df_match, df_bring = split_data(df_pre)
        print("=== 第三步：匹配/验证 ===")
        df_m_algo = match_partners(df_match)
        df_m_bring = verify_bring_partners(df_bring)
        df_all = pd.concat([df_m_algo, df_m_bring], ignore_index=True)
        df_all.to_excel('excel4_所有匹配结果.xlsx', index=False)
        print("=== 第四步：分配场次 ===")
        assign_dance_sessions(df_all, df_m_bring)
        print("全部完成！文件：excel1~5")
    except Exception as e:
        print("程序执行出错:", e)
        traceback.print_exc()


if __name__ == "__main__":
    main()
