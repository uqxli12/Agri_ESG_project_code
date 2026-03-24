import matplotlib.pyplot as plt
import numpy as np

# 数据
categories = [
    'PV Agriculture', 'Soil Analysis', 'Nutrient Recovery',
    'Gene Editing', 'New Fertilizers', 'Phenotype Monitoring'
]
esg_data = [
    [71.69,71.69,71.2,71.2,71.2,71.2,70.67,70.13,70.12,69.65,69.6,69.59,69.59,69.01,68.97,68.73,68.72,68.4,68.4,68,67.68,67.68,67.68,67.68,66.5],
    [66.48,66.48,65.4,64.83,64.79,63.86],
    [66.38,64.99,64.95,64.95,64.06,64.02],
    [64.81,64.62,63.91],
    [64.16,64.16,63.98],
    [66.14]
]

# CNS 绘图风格
plt.rcParams['font.family'] = 'Arial'
plt.rcParams['font.size'] = 9
plt.rcParams['axes.linewidth'] = 1.0
plt.rcParams['xtick.direction'] = 'in'
plt.rcParams['ytick.direction'] = 'in'

fig, ax = plt.subplots(figsize=(7,5), dpi=300)

box = ax.boxplot(
    esg_data, labels=categories, patch_artist=True,
    boxprops=dict(facecolor='#a1d9ff', alpha=0.8),
    medianprops=dict(color='#c92c6c', linewidth=1.5)
)
colors = ['#006999', '#0077b6', '#0096c7', '#00b4d8', '#48cae4', '#90e0ef']
for i, d in enumerate(esg_data):
    x = np.random.normal(i + 1, 0.06, len(d))
    ax.scatter(x, d, color=colors[i], s=15, alpha=0.8)

ax.set_title('A. ESG Distribution by Tech Category', fontsize=11, weight='bold')
ax.set_ylabel('ESG Score', fontsize=10, weight='bold')
ax.set_ylim(62, 73)
ax.grid(axis='y', linestyle='--', alpha=0.2)
ax.tick_params(axis='x', labelsize=7.5, rotation=30)
ax.tick_params(axis='y', labelsize=8)

plt.tight_layout()
plt.savefig('Figure_A_ESG_Distribution.png', dpi=600, bbox_inches='tight')
plt.show()