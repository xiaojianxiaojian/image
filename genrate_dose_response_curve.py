import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from scipy.optimize import curve_fit

# 从Excel读取数据
df = pd.read_excel('2.xlsx', header=None)

# ============ 数据清洗函数 ============
def remove_outliers_iqr(values, multiplier=1.5):
    """使用IQR方法去除离群值"""
    if len(values) < 4:
        return values
    q1 = np.percentile(values, 25)
    q3 = np.percentile(values, 75)
    iqr = q3 - q1
    lower = q1 - multiplier * iqr
    upper = q3 + multiplier * iqr
    return [v for v in values if lower <= v <= upper]

# ============ 第一部分：左侧 T+G 数据 (columns 0-5) ============
tg_columns = list(range(0, 6))
tg_headers = [str(df.iloc[0, col]) for col in tg_columns]

tg_data = {}
for col, header in zip(tg_columns, tg_headers):
    values = df.iloc[1:, col].dropna().astype(float).tolist()
    tg_data[header] = values

# 计算T+G数据的平均值和标准误差
tg_means = []
tg_sems = []
for header in tg_headers:
    values = tg_data[header]
    mean = np.mean(values)
    sem = np.std(values, ddof=1) / np.sqrt(len(values))
    tg_means.append(mean)
    tg_sems.append(sem)

tg_x_positions = [0.01, 0.05, 0.1, 2.5, 5, 10]

# ============ 第二部分：右侧 PAN 数据 (columns 8-13) ============
pan_columns = list(range(8, 14))
pan_headers = [str(df.iloc[0, col]) for col in pan_columns]

pan_data = {}
pan_data_cleaned = {}
for col, header in zip(pan_columns, pan_headers):
    values = df.iloc[1:, col].dropna().astype(float).tolist()
    pan_data[header] = values
    # 去除离群值
    cleaned = remove_outliers_iqr(values, multiplier=1.5)
    pan_data_cleaned[header] = cleaned
    if len(cleaned) < len(values):
        print(f'{header}: 去除 {len(values) - len(cleaned)} 个离群值')

# 计算PAN数据的平均值和标准误差（使用清洗后的数据）
pan_means = []
pan_sems = []
for header in pan_headers:
    values = pan_data_cleaned[header]
    mean = np.mean(values)
    sem = np.std(values, ddof=1) / np.sqrt(len(values))
    pan_means.append(mean)
    pan_sems.append(sem)

pan_x_positions = [0.01, 0.1, 1, 10, 100, 1000]

# ============ 4参数逻辑斯谛模型(EC50) ============
def four_parameter_logistic(x, bottom, top, logEC50, hillslope):
    """4-parameter logistic dose-response model"""
    return bottom + (top - bottom) / (1 + 10**((logEC50 - np.log10(x)) * hillslope))

# 拟合 T+G 数据
try:
    tg_p0 = [min(tg_means)*0.9, max(tg_means)*1.1, np.log10(0.5), 1.0]
    tg_popt, _ = curve_fit(four_parameter_logistic, tg_x_positions, tg_means, p0=tg_p0, maxfev=10000)
    print(f"T+G EC50拟合: bottom={tg_popt[0]:.2f}, top={tg_popt[1]:.2f}, EC50={10**tg_popt[2]:.2f}, hillslope={tg_popt[3]:.2f}")
except:
    tg_popt = [min(tg_means)*0.5, max(tg_means)*1.2, np.log10(0.5), 1.0]

# 拟合 PAN 数据 - 使用4PL模型
try:
    pan_p0 = [min(pan_means)*0.9, max(pan_means)*1.1, np.log10(10), 0.5]
    pan_popt, _ = curve_fit(four_parameter_logistic, pan_x_positions, pan_means, p0=pan_p0, maxfev=10000)
    print(f"PAN EC50拟合: bottom={pan_popt[0]:.2f}, top={pan_popt[1]:.2f}, EC50={10**pan_popt[2]:.2f}, hillslope={pan_popt[3]:.2f}")
except:
    pan_popt = [min(pan_means)*0.5, max(pan_means)*1.2, np.log10(10), 0.5]

# ============ 创建图形 - 双Y轴 ============
fig, ax1 = plt.subplots(figsize=(12, 8))
ax2 = ax1.twinx()  # 创建第二个Y轴

# T+G 数据 - 蓝色，使用左Y轴，带误差棒
ax1.errorbar(tg_x_positions, tg_means, yerr=tg_sems, fmt='o', color='blue',
             markersize=12, capsize=5, linewidth=2, ecolor='blue', label='T+G')

# PAN 数据 - 红色，使用右Y轴，带误差棒（误差棒缩小为50%）
pan_sems_scaled = [sem * 0.5 for sem in pan_sems]
ax2.errorbar(pan_x_positions, pan_means, yerr=pan_sems_scaled, fmt='s', color='red',
             markersize=12, capsize=5, linewidth=2, ecolor='red', label='PAN')

# 绘制T+G拟合曲线
tg_x_smooth = np.logspace(np.log10(0.01), np.log10(1000), 500)
tg_y_smooth = four_parameter_logistic(tg_x_smooth, *tg_popt)
ax1.plot(tg_x_smooth, tg_y_smooth, 'b-', linewidth=2)

# 绘制PAN拟合曲线
pan_x_smooth = np.logspace(np.log10(0.01), np.log10(1000), 500)
pan_y_smooth = four_parameter_logistic(pan_x_smooth, *pan_popt)
ax2.plot(pan_x_smooth, pan_y_smooth, 'r-', linewidth=2)

# 设置x轴为对数刻度
ax1.set_xscale('log')
ax1.set_xlim(0.005, 2000)

# 设置左Y轴 (T+G) 范围
ax1.set_ylim(0, 60)
ax1.set_ylabel('SiglecF + Neutrophils\n(% of Neutrophil) - T+G', fontsize=16, color='blue')
ax1.tick_params(axis='y', labelsize=14, colors='blue')

# 设置右Y轴 (PAN) 范围
ax2.set_ylim(4.2, 6.6)
ax2.set_ylabel('SiglecF + Neutrophils\n(% of Neutrophil) - PAN', fontsize=16, color='red')
ax2.tick_params(axis='y', labelsize=14, colors='red')

# 设置x轴刻度标签
ax1.set_xticks([0.01, 0.05, 0.1, 1, 10, 100, 1000])
ax1.set_xticklabels(['con', '0.05', '0.1', '1', '10', '100', '1K'])
ax1.tick_params(axis='x', labelsize=14, length=6, width=1.5)

# 设置标题
ax1.set_title('Dose-response curve', fontsize=24, fontweight='bold', pad=20)

# 设置网格
ax1.grid(True, which='both', linestyle=':', alpha=0.3, color='gray')

# 添加图例
lines1, labels1 = ax1.get_legend_handles_labels()
lines2, labels2 = ax2.get_legend_handles_labels()
ax1.legend(lines1 + lines2, labels1 + labels2, fontsize=14, loc='upper left')

# 调整布局
plt.tight_layout()

# 保存图片
plt.savefig('dose_response_curve_2xlsx.svg', format='svg', bbox_inches='tight')
print("图片已保存为: dose_response_curve_2xlsx.svg")

# 显示图片
plt.show()
