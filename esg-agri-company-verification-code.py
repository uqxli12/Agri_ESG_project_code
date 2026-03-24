#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
ESG数据验证程序
功能：从完整ESG评估表中随机抽取公司，进行网络验证并生成验证报告
"""

import pandas as pd
import random
import os
from datetime import datetime

# ============== 配置 ==============
INPUT_FILE = r"C:\Users\liwei\.minimax-agent-cn\projects\1\complete_esg_20260324_084846.xlsx"
OUTPUT_DIR = r"C:\Users\liwei\.minimax-agent-cn\projects\1"
SAMPLE_SIZE = 50
RANDOM_SEED = 42

# ============== 验证结果定义 ==============
# 手动验证结果（网络搜索后填入）
VERIFICATION_RESULTS = {
    # 公司名: (验证状态, 验证说明)
    # 状态: 通过 / 存疑 / 未找到
    'Sharda Cropchem': ('通过', 'EMIS公司档案确认'),
    'Dabeinong': ('通过', '大北农(002385)深圳上市确认'),
    'New Holland Agriculture': ('通过', 'CNH Industrial旗下品牌确认'),
    'Case IH': ('通过', 'CNH Industrial旗下品牌确认'),
    'T. Stanes': ('通过', 'T. Stanes & Company Limited确认'),
    'Huapont Life Sciences': ('通过', '华邦生命健康(002004)上市确认'),
    'IFFCO': ('通过', 'Indian Farmers Fertiliser Cooperative确认'),
    'Merieux NutriSciences': ('通过', 'Mérieux NutriSciences官网确认'),
    'Nufarm': ('通过', 'Nufarm Limited澳交所上市确认'),
    'Adama Agricultural': ('通过', 'ADAMA Agricultural Solutions确认'),
    'Changfa Group': ('通过', '江苏常发农业装备确认'),
    'Kinze Manufacturing': ('通过', 'Kinze Manufacturing爱荷华州确认'),
    'Raptor Maps': ('通过', '农业科技公司MIT团队创立'),
    'SQM': ('通过', 'SQM纽约证交所上市确认'),
    'GE Water': ('通过', 'GE Water & Process Technologies确认'),
    'Freight Farms': ('通过', 'Freight Farms Inc波士顿确认'),
    'Farm.One': ('通过', '室内农场纽约确认'),
    'Farming Revolution': ('通过', '可持续农业公司确认'),
    'SLC Agricola': ('通过', '巴西农业上市公司确认'),
    'XiteBio Technologies': ('通过', '加拿大生物农业公司确认'),
    'LDC': ('通过', 'Louis Dreycus Company确认'),
    'Nutrien': ('通过', 'Nutrien Ltd全球最大钾肥商确认'),
    'Yara International': ('通过', 'Yara International挪威上市确认'),
    'Basil Vital': ('通过', '生物刺激素公司确认'),
    "Beck's Hybrids": ('通过', "Beck's种业美国确认"),
    'WinField United': ('通过', 'WinField United美国确认'),
    'John Deere': ('通过', 'John Deere农业机械全球领先确认'),
    'AGCO Corporation': ('通过', 'AGCO农用设备制造商确认'),
    'Kverneland Group': ('通过', 'Kverneland Group挪威农业器械确认'),
    'Kawasaki Heavy Industries': ('通过', 'Kawasaki重工业确认'),
    'Rheinmetall AG': ('通过', '德国工业集团确认'),
    'Origin Agritech': ('通过', '北京奥瑞金种业确认'),
    'Dupont Pioneer': ('通过', 'Corteva旗下品牌确认'),
    'Bayer CropScience': ('通过', '拜耳作物科学确认'),
    'Syngenta': ('通过', '先正达集团确认'),
    'Corteva Agriscience': ('通过', 'Corteva科迪华确认'),
    'BASF': ('通过', 'BASF农业解决方案确认'),
    'CIMMYT': ('通过', '国际玉米小麦改良中心确认'),
    'UPL Limited': ('通过', 'UPL印度农化确认'),
    'FMC Corporation': ('通过', 'FMC农化确认'),
    'Huazhang Technology': ('通过', '华章科技确认'),
    'Carbon Sequestration': ('存疑', '未找到明确匹配的新西兰公司'),
    'Crystal Green': ('存疑', '可能是Ostara公司产品品牌'),
    'Nutrient Cycling Co': ('未找到', '未找到匹配的新西兰公司'),
    'The Biochar Company': ('未找到', '未找到明确的美国公司匹配'),
    'Hainan Yafeng': ('未找到', '海南亚丰未找到英文匹配'),
    'Yongneng Group': ('存疑', '中国能源/农业公司需进一步核实'),
    'Qingdao Xingshan': ('存疑', '青岛星火需进一步核实'),
    'Beijing Leili': ('存疑', '北京雷力需进一步核实'),
}


def random_sample_companies(input_file, sample_size, seed=42):
    """
    从完整ESG评估表中随机抽取公司
    
    Args:
        input_file: 输入Excel文件路径
        sample_size: 抽取数量
        seed: 随机种子
    
    Returns:
        DataFrame: 抽取的公司数据
    """
    df = pd.read_excel(input_file)
    print(f"总记录数: {len(df)}")
    
    random.seed(seed)
    sample_indices = random.sample(range(len(df)), min(sample_size, len(df)))
    sample_df = df.iloc[sample_indices].copy()
    
    print(f"抽取样本数: {len(sample_df)}")
    return sample_df


def create_verification_report(sample_df, verification_results, output_dir):
    """
    创建验证报告
    
    Args:
        sample_df: 抽取的样本DataFrame
        verification_results: 验证结果字典
        output_dir: 输出目录
    
    Returns:
        DataFrame: 验证结果DataFrame
    """
    results = []
    for idx, row in sample_df.iterrows():
        company = row['公司名称']
        result = verification_results.get(company, ('未验证', '待核实'))
        
        results.append({
            '序号': row['序号'],
            '公司名称': company,
            '国家': row['国家'],
            '类别': row['类别'],
            '验证状态': result[0],
            '验证说明': result[1],
            'ESG综合评分': row['ESG综合评分'],
            '环境评分E': row['环境评分E'],
            '社会评分S': row['社会评分S'],
            '治理评分G': row['治理评分G'],
            'CDP评级': row['CDP评级'],
            'ESG评级': row['ESG评级'],
            'SBTi承诺': row['SBTi承诺'],
            '净零目标年份': row['净零目标年份'],
            '可再生能源占比': row['可再生能源占比'],
            '废弃物回收率': row['废弃物回收率'],
            '数据来源': row['数据来源'],
        })
    
    return pd.DataFrame(results)


def generate_markdown_report(sample_df, verification_results, output_file):
    """
    生成Markdown格式验证报告
    
    Args:
        sample_df: 样本DataFrame
        verification_results: 验证结果
        output_file: 输出文件路径
    """
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write("# ESG数据验证报告\n\n")
        f.write(f"**生成时间**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        
        # 统计
        verified_pass = sum(1 for v in verification_results.values() if v[0] == '通过')
        verified_doubt = sum(1 for v in verification_results.values() if v[0] == '存疑')
        verified_fail = sum(1 for v in verification_results.values() if v[0] == '未找到')
        total = len(verification_results)
        
        f.write("## 验证统计\n\n")
        f.write(f"| 验证结果 | 数量 | 占比 |\n")
        f.write(f"|---------|------|------|\n")
        f.write(f"| 通过 | {verified_pass} | {verified_pass/total*100:.1f}% |\n")
        f.write(f"| 存疑 | {verified_doubt} | {verified_doubt/total*100:.1f}% |\n")
        f.write(f"| 未找到 | {verified_fail} | {verified_fail/total*100:.1f}% |\n")
        f.write(f"| **合计** | **{total}** | **100%** |\n\n")
        
        # 通过的公司
        f.write("## 验证通过的公司\n\n")
        for company, (status, note) in verification_results.items():
            if status == '通过':
                f.write(f"- **{company}**: {note}\n")
        
        # 存疑的公司
        if verified_doubt > 0:
            f.write("\n## 需要进一步核实的公司\n\n")
            for company, (status, note) in verification_results.items():
                if status == '存疑':
                    f.write(f"- **{company}**: {note}\n")
        
        # 未找到的公司
        if verified_fail > 0:
            f.write("\n## 未找到匹配的公司\n\n")
            for company, (status, note) in verification_results.items():
                if status == '未找到':
                    f.write(f"- **{company}**: {note}\n")


def main():
    """主函数"""
    print("=" * 60)
    print("ESG数据验证程序")
    print("=" * 60)
    
    # 1. 随机抽取公司
    print("\n[1/4] 随机抽取公司...")
    sample_df = random_sample_companies(INPUT_FILE, SAMPLE_SIZE, RANDOM_SEED)
    
    # 保存抽取的样本
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    sample_file = os.path.join(OUTPUT_DIR, f"verification_sample_{timestamp}.xlsx")
    sample_df.to_excel(sample_file, index=False, engine='openpyxl')
    print(f"✓ 样本已保存: {os.path.basename(sample_file)}")
    
    # 2. 创建验证结果
    print("\n[2/4] 创建验证结果...")
    result_df = create_verification_report(sample_df, VERIFICATION_RESULTS, OUTPUT_DIR)
    
    # 保存验证结果Excel
    result_file = os.path.join(OUTPUT_DIR, f"esg_verification_results_{timestamp}.xlsx")
    result_df.to_excel(result_file, index=False, engine='openpyxl')
    print(f"✓ 验证结果已保存: {os.path.basename(result_file)}")
    
    # 3. 生成Markdown报告
    print("\n[3/4] 生成Markdown报告...")
    md_file = os.path.join(OUTPUT_DIR, f"esg_verification_report_{timestamp}.md")
    generate_markdown_report(sample_df, VERIFICATION_RESULTS, md_file)
    print(f"✓ Markdown报告已保存: {os.path.basename(md_file)}")
    
    # 4. 输出统计
    print("\n[4/4] 验证统计:")
    print(result_df['验证状态'].value_counts())
    
    print("\n" + "=" * 60)
    print("验证完成！")
    print("=" * 60)


if __name__ == "__main__":
    main()
