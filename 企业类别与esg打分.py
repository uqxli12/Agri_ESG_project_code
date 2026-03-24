import matplotlib.pyplot as plt
import seaborn as sns
import pandas as pd
import numpy as np

# 数据整理（直接内置在代码中）
data = {
    "Category": ["PV Agriculture"]*25 + ["Soil Analysis"]*6 + ["Nutrient Recovery"]*6 + ["Phenotype Monitoring"]*1 + ["Gene Editing"]*3 + ["New Fertilizers"]*3,
    "ESG": [
        71.69,71.69,71.2,71.2,71.2,71.2,70.67,70.13,70.12,69.65,69.6,69.59,69.59,69.01,68.97,68.73,68.72,68.4,68.4,68,67.68,67.68,67.68,67.68,66.5,
        66.48,66.48,65.4,64.83,64.79,63.86,
        66.38,64.99,64.95,64.95,64.06,64.02,
        66.14,
        64.81,64.62,63.91,
        64.16,64.16,63.98
    ]
}
df = pd.DataFrame(data)

# 绘图风格
plt.rcParams['font.family'] = 'Arial'
plt.figure(figsize=(10,6))

# 绘制箱线+散点图
sns.boxplot(x="Category", y="ESG", data=df, palette="YlGnBu", showmeans=True, meanprops={"marker":"o","markerfacecolor":"red"})
sns.stripplot(x="Category", y="ESG", data=df, color="black", alpha=0.6, size=4)

# 标签
plt.title('ESG Performance of Tech-Enabled Agriculture Sectors', fontsize=14, weight='bold')
plt.xlabel('Agricultural Technology Category', fontsize=12)
plt.ylabel('ESG Comprehensive Score', fontsize=12)
plt.xticks(rotation=15, fontsize=10)
plt.tight_layout()
plt.savefig('agri_tech_esg_comparison.png', dpi=300)
plt.show()