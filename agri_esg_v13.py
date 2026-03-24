#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
全球农业企业ESG调研分析系统 v13.0
修复版CSV解析 - 正确处理两种数据格式

格式A (8字段): 序号,公司名，国家，类别，CDP/SBTi，10指标，小农户，3社评，数据
格式B (8字段): 序号,公司名(含逗号),国家，类别，CDP/SBTi，10指标，小农户，3社评，数据
"""

import pandas as pd
import numpy as np
from datetime import datetime
import os
import re
import warnings
warnings.filterwarnings('ignore')

pd.set_option('display.unicode.ambiguous_as_wide', True)
pd.set_option('display.unicode.east_asian_width', True)

OUTPUT_DIR = r"C:\Users\liwei\.minimax-agent-cn\projects\1"
os.makedirs(OUTPUT_DIR, exist_ok=True)

COLUMNS = [
    '序号', '公司名称', '国家', '类别', 
    'CDP评级', 'SBTi承诺', '净零目标年份', '可再生能源占比',
    '碳强度', '水使用强度', '废弃物回收率', '绿色专利占比',
    '研发支出占比', '可持续产品营收占比', 'ESG评级', '争议事件数量',
    '小农户支持', '供应链审核率', '董事会多样性', 'ESG薪酬挂钩',
    '数据来源', '数据质量评分'
]

COUNTRIES = ['美国', '德国', '中国', '法国', '丹麦', '荷兰', '日本', '韩国', '印度', 
             '巴西', '瑞士', '英国', '澳大利亚', '加拿大', '挪威', '意大利', '西班牙',
             '比利时', '以色列', '瑞典', '芬兰', '奥地利', '新加坡', '南非', '新西兰',
             '泰国', '捷克', '香港', '俄罗斯', '智利', '卢森堡', '瑞士/中国', '美国/意大利']

def parse_row_v13(line):
    """v13解析单行数据 - 正确处理两种格式"""
    line = line.strip()
    if not line:
        return None
    
    parts = line.split('，')
    row = [''] * 22
    
    first_field = parts[0] if parts else ''
    
    # 提取序号
    seq_match = re.match(r'^(\d+)', first_field)
    if not seq_match:
        return None
    row[0] = seq_match.group(1)
    
    # 分析第一字段的逗号数量
    first_parts = first_field.split(',')
    
    if len(first_parts) == 2:
        # ========== 格式A ==========
        # parts: [序号+公司, 国家, 类别, CDP/SBTi, 10指标, 小农户, 3社评, 数据]
        row[1] = first_parts[1].strip()  # 公司名
        row[2] = parts[1] if len(parts) > 1 else ''  # 国家
        row[3] = parts[2] if len(parts) > 2 else ''  # 类别
        
        # CDP/SBTi
        if len(parts) > 3:
            cdp_sbti = parts[3].split(',')
            row[4] = cdp_sbti[0].strip()  # CDP评级
            row[5] = cdp_sbti[1].strip() if len(cdp_sbti) > 1 else ''  # SBTi承诺
        
        # 10个指标
        if len(parts) > 4:
            metrics = parts[4].split(',')
            if len(metrics) >= 10:
                row[6] = metrics[0]   # 净零目标年份
                row[7] = metrics[1]    # 可再生能源占比
                row[8] = metrics[2]    # 碳强度
                row[9] = metrics[3]   # 水使用强度
                row[10] = metrics[4]  # 废弃物回收率
                row[11] = metrics[5]  # 绿色专利占比
                row[12] = metrics[6]  # 研发支出占比
                row[13] = metrics[7]  # 可持续产品营收占比
                row[14] = metrics[8]  # ESG评级
                row[15] = metrics[9]  # 争议事件数量
        
        # 小农户支持
        if len(parts) > 5:
            row[16] = parts[5]
        
        # 3个社会指标
        if len(parts) > 6:
            social = parts[6].split(',')
            row[17] = social[0].strip() if len(social) > 0 else ''  # 供应链审核率
            row[18] = social[1].strip() if len(social) > 1 else ''  # 董事会多样性
            row[19] = social[2].strip() if len(social) > 2 else ''  # ESG薪酬挂钩
        
        # 数据来源和评分
        if len(parts) > 7:
            data_parts = parts[7:]
            full_text = '，'.join(data_parts)
            score_match = re.search(r'[，,](\d+)$', full_text)
            if score_match:
                row[21] = score_match.group(1)
                row[20] = full_text[:score_match.start()].rstrip('，').strip()
            else:
                row[20] = full_text
    else:
        # ========== 格式B ==========
        # parts: [序号+公司+国家, 类别, CDP/SBTi, 10指标, 小农户, 3社评, 数据]
        # 公司名中包含英文逗号，国家被合并到第一个字段
        
        # 查找国家
        country_idx = None
        for i in range(1, len(first_parts)):
            for c in COUNTRIES:
                if c in first_parts[i]:
                    country_idx = i
                    row[2] = c
                    row[1] = ','.join(first_parts[1:country_idx])  # 公司名
                    break
            if country_idx:
                break
        
        if not country_idx:
            row[1] = first_parts[1].strip() if len(first_parts) > 1 else ''
        
        row[3] = parts[1] if len(parts) > 1 else ''  # 类别
        
        # CDP/SBTi
        if len(parts) > 2:
            cdp_sbti = parts[2].split(',')
            row[4] = cdp_sbti[0].strip()  # CDP评级
            row[5] = cdp_sbti[1].strip() if len(cdp_sbti) > 1 else ''  # SBTi承诺
        
        # 10个指标
        if len(parts) > 3:
            metrics = parts[3].split(',')
            if len(metrics) >= 10:
                row[6] = metrics[0]   # 净零目标年份
                row[7] = metrics[1]    # 可再生能源占比
                row[8] = metrics[2]    # 碳强度
                row[9] = metrics[3]   # 水使用强度
                row[10] = metrics[4]  # 废弃物回收率
                row[11] = metrics[5]  # 绿色专利占比
                row[12] = metrics[6]  # 研发支出占比
                row[13] = metrics[7]  # 可持续产品营收占比
                row[14] = metrics[8]  # ESG评级
                row[15] = metrics[9]  # 争议事件数量
        
        # 小农户支持
        if len(parts) > 4:
            row[16] = parts[4]
        
        # 3个社会指标
        if len(parts) > 5:
            social = parts[5].split(',')
            row[17] = social[0].strip() if len(social) > 0 else ''  # 供应链审核率
            row[18] = social[1].strip() if len(social) > 1 else ''  # 董事会多样性
            row[19] = social[2].strip() if len(social) > 2 else ''  # ESG薪酬挂钩
        
        # 数据来源和评分
        if len(parts) > 6:
            data_parts = parts[6:]
            full_text = '，'.join(data_parts)
            score_match = re.search(r'[，,](\d+)$', full_text)
            if score_match:
                row[21] = score_match.group(1)
                row[20] = full_text[:score_match.start()].rstrip('，').strip()
            else:
                row[20] = full_text
    
    return row

def parse_csv_v13(filepath):
    """v13 CSV解析"""
    with open(filepath, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    data = []
    for line in lines[1:]:
        row = parse_row_v13(line)
        if row:
            data.append(row)
    
    return pd.DataFrame(data, columns=COLUMNS)

def calculate_esg_score(row):
    """计算ESG综合评分"""
    weights = {'environmental': 0.50, 'social': 0.30, 'governance': 0.20}
    
    def score_value(val, default=50):
        if pd.isna(val) or val == '' or val == '-' or val is None:
            return default
        val_str = str(val).strip()
        if not val_str:
            return default
        
        cdp_map = {'A': 100, 'A-': 95, 'B': 85, 'B-': 75, 'C': 60, 'C-': 50, 'D': 40}
        if val_str in cdp_map:
            return cdp_map[val_str]
        
        esg_map = {'AAA': 100, 'AA': 95, 'A': 85, 'BBB': 75, 'BB': 60, 'B': 45, 'CCC': 30}
        if val_str in esg_map:
            return esg_map[val_str]
        
        sbti_map = {'已承诺': 100, '在评估中': 60, '未承诺': 20}
        if val_str in sbti_map:
            return sbti_map[val_str]
        
        if val_str in ['有', '是']:
            return 100
        if val_str in ['无', '否']:
            return 20
        if val_str == '部分':
            return 60
        
        try:
            num = float(val_str.replace('%', '').strip())
            return min(100, max(0, num))
        except:
            return default
    
    e1 = score_value(row.get('CDP评级', '')) * 3.0
    e2 = score_value(row.get('SBTi承诺', '')) * 2.0
    e3 = score_value(row.get('净零目标年份', ''), 30) * 1.5
    e4 = score_value(row.get('可再生能源占比', '')) * 2.5
    e5 = score_value(row.get('碳强度', ''), 70) * 2.0
    e6 = score_value(row.get('水使用强度', ''), 70) * 1.5
    e7 = score_value(row.get('废弃物回收率', '')) * 2.0
    e8 = score_value(row.get('绿色专利占比', '')) * 1.5
    e_final = (e1 + e2 + e3 + e4 + e5 + e6 + e7 + e8) / 15.5
    
    s1 = score_value(row.get('研发支出占比', '')) * 2.0
    s2 = score_value(row.get('可持续产品营收占比', '')) * 2.5
    s3 = score_value(row.get('小农户支持', '')) * 1.5
    s4 = score_value(row.get('供应链审核率', '')) * 2.0
    s5 = score_value(row.get('董事会多样性', '')) * 1.5
    s6 = score_value(row.get('ESG薪酬挂钩', '')) * 1.5
    s_final = (s1 + s2 + s3 + s4 + s5 + s6) / 11.0
    
    g1 = score_value(row.get('ESG评级', '')) * 3.0
    g2 = score_value(row.get('争议事件数量', ''), 70) * 2.5
    g3 = score_value(row.get('数据质量评分', '')) * 2.0
    g4 = score_value(row.get('数据来源', '')) * 2.5
    g_final = (g1 + g2 + g3 + g4) / 10.0
    
    esg = e_final * weights['environmental'] + s_final * weights['social'] + g_final * weights['governance']
    
    return pd.Series({
        'ESG综合评分': round(esg, 2),
        '环境评分E': round(e_final, 2),
        '社会评分S': round(s_final, 2),
        '治理评分G': round(g_final, 2)
    })

def main():
    print("=" * 60)
    print("全球农业企业ESG调研分析系统 v13.0")
    print("=" * 60)
    
    print("\n[1/5] 加载企业数据...")
    filepath = r"C:\Users\liwei\.minimax-agent-cn\projects\1\agri-company-metrics.txt"
    df = parse_csv_v13(filepath)
    print(f"✓ 加载 {len(df)} 家企业")
    
    print(f"\n  数据验证(第1-5行 格式A):")
    for i, row in df.head(5).iterrows():
        print(f"  [{i+1}] 序号:{row['序号']} | 公司:{row['公司名称']} | 国家:{row['国家']} | 类别:{row['类别']} | CDP:{row['CDP评级']} | 净零:{row['净零目标年份']} | 可再生能源:{row['可再生能源占比']}")
    
    print(f"\n  数据验证(第9-10行 格式B):")
    for i in [8, 9]:
        if i < len(df):
            row = df.iloc[i]
            print(f"  [{i+1}] 序号:{row['序号']} | 公司:{row['公司名称']} | 国家:{row['国家']} | 类别:{row['类别']} | CDP:{row['CDP评级']} | 净零:{row['净零目标年份']} | 可再生能源:{row['可再生能源占比']}")
    
    print(f"\n  解析统计:")
    print(f"  - 唯一公司数: {df['公司名称'].nunique()}")
    print(f"  - 唯一类别数: {df['类别'].nunique()}")
    print(f"  - 唯一国家数: {df['国家'].nunique()}")
    print(f"  - CDP评级分布: {df['CDP评级'].value_counts().head(5).to_dict()}")
    print(f"  - 空国家数: {(df['国家'] == '').sum()}")
    
    print("\n[2/5] 计算ESG评分...")
    scores = df.apply(calculate_esg_score, axis=1)
    df = pd.concat([df, scores], axis=1)
    print(f"✓ 计算完成，平均ESG: {df['ESG综合评分'].mean():.2f}")
    
    print("\n[3/5] 按类别分析...")
    category_stats = df.groupby('类别').agg({
        'ESG综合评分': 'mean', '公司名称': 'count'
    }).rename(columns={'公司名称': '企业数'})
    category_stats = category_stats.sort_values('ESG综合评分', ascending=False)
    print(f"\n共 {len(category_stats)} 个类别:")
    print(category_stats.head(10).round(2).to_string())
    
    print("\n[4/5] 按国家分析...")
    country_stats = df.groupby('国家').agg({
        'ESG综合评分': 'mean', '公司名称': 'count'
    }).rename(columns={'公司名称': '企业数'})
    country_stats = country_stats.sort_values('ESG综合评分', ascending=False)
    print(f"\n共 {len(country_stats)} 个国家:")
    print(country_stats.head(10).round(2).to_string())
    
    print("\n[5/5] 导出结果...")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    output_cols = ['序号', '公司名称', '国家', '类别', 'ESG综合评分', '环境评分E', '社会评分S', '治理评分G'] + COLUMNS[4:]
    complete_df = df[[c for c in output_cols if c in df.columns]].copy()
    complete_df.to_excel(os.path.join(OUTPUT_DIR, f"complete_esg_{timestamp}.xlsx"), index=False)
    category_stats.to_excel(os.path.join(OUTPUT_DIR, f"esg_by_category_{timestamp}.xlsx"))
    country_stats.to_excel(os.path.join(OUTPUT_DIR, f"esg_by_country_{timestamp}.xlsx"))
    df.nlargest(50, 'ESG综合评分')[['序号', '公司名称', '国家', '类别', 'ESG综合评分']].to_excel(
        os.path.join(OUTPUT_DIR, f"top50_esg_leaders_{timestamp}.xlsx"), index=False)
    
    print(f"✓ 已导出 4 个文件")
    
    print("\n" + "=" * 60)
    print("ESG领先企业 TOP 10")
    print("=" * 60)
    print(df.nlargest(10, 'ESG综合评分')[['公司名称', '国家', '类别', 'ESG综合评分']].round(2).to_string(index=False))
    
    print(f"\n📊 摘要: 企业{len(df)}家 | 国家{df['国家'].nunique()}个 | 类别{df['类别'].nunique()}类 | 平均ESG:{df['ESG综合评分'].mean():.2f}")

if __name__ == "__main__":
    main()
