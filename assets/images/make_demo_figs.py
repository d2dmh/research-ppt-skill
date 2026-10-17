"""演示稿用的具体数据图（示意数据）：园区综合能源系统负荷预测与优化调度。"""
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib import font_manager
from pathlib import Path

out = Path(__file__).parent
out.mkdir(exist_ok=True)
for f in font_manager.findSystemFonts():
    if 'NotoSansCJK-Regular' in f or 'NotoSansCJK-Medium' in f:
        font_manager.fontManager.addfont(f)
plt.rcParams.update({'font.family': 'Noto Sans CJK JP', 'font.size': 13, 'axes.spines.top': False,
                     'axes.spines.right': False, 'axes.edgecolor': '#55606F', 'axes.linewidth': 1.0,
                     'axes.titlesize': 14, 'axes.titleweight': 'bold', 'axes.titlecolor': '#073DA5',
                     'legend.frameon': False, 'savefig.dpi': 200, 'axes.unicode_minus': False})
NAVY, BLUE, LB, RED, GREY, ORANGE = '#073DA5', '#3E73CB', '#9DB8E8', '#DE1515', '#8A94A6', '#E39B2D'
rng = np.random.default_rng(7)
h = np.arange(24)


def save(fig, name):
    fig.tight_layout()
    fig.savefig(out / f'{name}.png', facecolor='white')
    plt.close(fig)


# 1 典型日电 / 热 / 冷负荷
fig, ax = plt.subplots(figsize=(6.4, 3.6))
elec = 3.2 + 1.6 * np.exp(-((h - 10) / 3.2) ** 2) + 1.3 * np.exp(-((h - 19) / 2.4) ** 2)
heat = 2.4 + 1.2 * np.cos((h - 6) / 24 * 2 * np.pi)
cool = 0.6 + 2.0 * np.exp(-((h - 15) / 3.5) ** 2)
ax.plot(h, elec, color=NAVY, lw=2.4, label='电负荷')
ax.plot(h, heat, color=RED, lw=2.0, label='热负荷')
ax.plot(h, cool, color=BLUE, lw=2.0, ls='--', label='冷负荷')
ax.set_xlabel('时刻 / h'); ax.set_ylabel('功率 / MW'); ax.set_xticks(range(0, 25, 4))
ax.legend(ncol=3, loc='upper left'); ax.set_title('园区典型日多能负荷（示意数据）')
save(fig, 'load')

# 2 光伏出力与分时电价
fig, ax = plt.subplots(figsize=(6.4, 3.6))
pv = np.clip(4.2 * np.sin((h - 6) / 12 * np.pi), 0, None)
price = np.where((h >= 8) & (h < 11) | (h >= 18) & (h < 22), 1.12, np.where((h >= 23) | (h < 7), 0.32, 0.68))
ax.fill_between(h, pv, color=LB, alpha=0.9, label='光伏出力')
ax.set_ylabel('出力 / MW'); ax.set_xlabel('时刻 / h'); ax.set_xticks(range(0, 25, 4))
ax2 = ax.twinx(); ax2.step(h, price, where='post', color=RED, lw=2.2, label='分时电价')
ax2.set_ylabel('电价 / (元/kWh)'); ax2.spines['top'].set_visible(False)
ax.set_title('光伏出力与分时电价错峰（示意数据）')
fig.legend(loc='upper right', bbox_to_anchor=(0.86, 0.86))
save(fig, 'pv_price')

# 3 预测 vs 实测
fig, ax = plt.subplots(figsize=(7.2, 3.6))
t = np.arange(96) / 4
real = 3.2 + 1.6 * np.exp(-((t - 10) / 3.2) ** 2) + 1.3 * np.exp(-((t - 19) / 2.4) ** 2) + rng.normal(0, 0.12, 96)
pred = real + rng.normal(0, 0.1, 96)
base = real + rng.normal(0, 0.32, 96) + 0.2 * np.sin(t / 3)
ax.fill_between(t, pred - 0.25, pred + 0.25, color=LB, alpha=0.6, label='90% 预测区间')
ax.plot(t, real, color='#222', lw=1.8, label='实测')
ax.plot(t, pred, color=RED, lw=1.8, label='本文方法')
ax.plot(t, base, color=GREY, lw=1.2, ls='--', label='基准 LSTM')
ax.set_xlabel('时刻 / h'); ax.set_ylabel('负荷 / MW'); ax.set_xticks(range(0, 25, 4))
ax.legend(ncol=4, loc='upper left', fontsize=11); ax.set_title('日前 15 min 负荷预测（示意数据）')
save(fig, 'forecast')

# 4 模型误差对比
fig, ax = plt.subplots(figsize=(5.6, 3.6))
models = ['ARIMA', 'XGBoost', 'LSTM', 'Transformer', '本文方法']
mape = [7.8, 5.6, 4.9, 4.3, 3.1]
bars = ax.bar(models, mape, color=[LB, LB, LB, BLUE, RED], width=0.6)
for b, v in zip(bars, mape):
    ax.text(b.get_x() + b.get_width() / 2, v + 0.15, f'{v}', ha='center', fontsize=12)
ax.set_ylabel('MAPE / %'); ax.set_title('不同模型预测误差（示意数据）'); ax.set_ylim(0, 9)
save(fig, 'mape')

# 5 负荷热力图（日 × 小时）
fig, ax = plt.subplots(figsize=(6.4, 3.6))
days = np.arange(60)
grid = (3 + 1.5 * np.exp(-((h[None, :] - 10) / 3.2) ** 2) + 1.2 * np.exp(-((h[None, :] - 19) / 2.4) ** 2)
        + 0.6 * np.sin(days[:, None] / 60 * 2 * np.pi) + 0.5 * ((days[:, None] % 7) < 5) + rng.normal(0, 0.12, (60, 24)))
im = ax.imshow(grid.T, aspect='auto', cmap='Blues', origin='lower')
ax.set_xlabel('日序'); ax.set_ylabel('时刻 / h'); ax.set_yticks(range(0, 24, 6))
fig.colorbar(im, ax=ax, label='MW', fraction=0.04)
ax.set_title('60 天负荷时空分布（示意数据）')
save(fig, 'heatmap')

# 6 帕累托前沿
fig, ax = plt.subplots(figsize=(5.6, 3.8))
c = np.linspace(0, 1, 40)
cost = 82 + 30 * c
carbon = 18 + 22 * np.exp(-3.2 * c)
ax.scatter(cost + rng.normal(0, 1.2, 40), carbon + rng.normal(0, 0.6, 40) + 3, s=14, color=LB, label='可行解')
ax.plot(cost, carbon, color=NAVY, lw=2.4, label='帕累托前沿')
ax.scatter([94], [18 + 22 * np.exp(-3.2 * 0.4)], s=110, color=RED, zorder=5, label='折中方案')
ax.set_xlabel('日运行成本 / 万元'); ax.set_ylabel('碳排放 / t'); ax.legend(loc='upper right')
ax.set_title('成本—碳排多目标权衡（示意数据）')
save(fig, 'pareto')

# 7 日前调度堆叠图
fig, ax = plt.subplots(figsize=(7.2, 3.6))
grid_buy = np.clip(elec - pv * 0.8 - 0.6, 0.2, None)
chp = np.full(24, 1.1)
ess = np.where((h >= 18) & (h < 22), 0.9, 0)
ax.stackplot(h, pv * 0.8, chp, ess, grid_buy, colors=[LB, BLUE, RED, '#C9D3E2'],
             labels=['光伏', '燃气轮机', '储能放电', '电网购电'])
ax.plot(h, elec + ess * 0, color='#222', lw=1.6, label='电负荷')
ax.set_xlabel('时刻 / h'); ax.set_ylabel('功率 / MW'); ax.set_xticks(range(0, 25, 4))
ax.legend(ncol=5, loc='upper left', fontsize=10); ax.set_title('日前优化调度结果（示意数据）')
save(fig, 'dispatch')

# 8 储能 SOC
fig, ax = plt.subplots(figsize=(5.6, 3.4))
soc = np.clip(0.3 + np.cumsum(np.where((h >= 0) & (h < 7), 0.06, 0) + np.where((h >= 11) & (h < 15), 0.05, 0)
                              - np.where((h >= 18) & (h < 22), 0.13, 0)), 0.1, 0.95)
ax.plot(h, soc * 100, color=NAVY, lw=2.4, marker='o', ms=4)
ax.axhspan(10, 90, color=LB, alpha=0.25)
ax.set_xlabel('时刻 / h'); ax.set_ylabel('SOC / %'); ax.set_xticks(range(0, 25, 4))
ax.set_title('储能荷电状态（示意数据）')
save(fig, 'soc')

# 9 预测散点
fig, ax = plt.subplots(figsize=(4.4, 3.8))
x = rng.uniform(2.5, 6.5, 300)
ax.scatter(x, x + rng.normal(0, 0.14, 300), s=8, color=BLUE, alpha=0.6)
ax.plot([2.5, 6.5], [2.5, 6.5], color=RED, lw=1.6)
ax.set_xlabel('实测 / MW'); ax.set_ylabel('预测 / MW'); ax.set_title('预测—实测一致性（示意数据）')
ax.text(2.7, 6.1, 'R² = 0.97', fontsize=13, color=NAVY, weight='bold')
save(fig, 'scatter')

# 10 稳健性箱线图
fig, ax = plt.subplots(figsize=(5.6, 3.6))
data = [rng.normal(m, s, 50) for m, s in [(3.1, 0.3), (3.4, 0.4), (3.6, 0.45), (4.2, 0.6)]]
bp = ax.boxplot(data, patch_artist=True, widths=0.5, medianprops={'color': RED, 'lw': 2})
for p in bp['boxes']:
    p.set_facecolor(LB)
ax.set_xticks([1, 2, 3, 4], ['正常日', '节假日', '极端天气', '数据缺失 20%'])
ax.set_ylabel('MAPE / %'); ax.set_title('不同工况下的误差分布（示意数据）')
save(fig, 'robust')

# 11 特征相关性
fig, ax = plt.subplots(figsize=(4.8, 4.0))
names = ['温度', '湿度', '辐照', '电价', '工作日', '负荷']
m = np.array([[1, -.3, .7, .2, 0, .6], [-.3, 1, -.4, 0, 0, -.2], [.7, -.4, 1, -.1, 0, .3],
              [.2, 0, -.1, 1, .3, .5], [0, 0, 0, .3, 1, .4], [.6, -.2, .3, .5, .4, 1]])
im = ax.imshow(m, cmap='RdBu_r', vmin=-1, vmax=1)
ax.set_xticks(range(6), names, rotation=45); ax.set_yticks(range(6), names)
for i in range(6):
    for j in range(6):
        ax.text(j, i, f'{m[i, j]:.1f}', ha='center', va='center', fontsize=10)
ax.set_title('特征相关性（示意数据）')
save(fig, 'corr')

# 12 碳排与成本年度对比
fig, ax = plt.subplots(figsize=(5.6, 3.6))
cats = ['运行成本', '碳排放', '弃光率', '峰谷差']
before = [100, 100, 100, 100]
after = [87.6, 78.4, 41.0, 72.3]
xx = np.arange(4)
ax.bar(xx - 0.18, before, 0.36, color='#C9D3E2', label='优化前')
ax.bar(xx + 0.18, after, 0.36, color=[NAVY, NAVY, RED, NAVY], label='优化后')
ax.set_xticks(xx, cats); ax.set_ylabel('相对值 / %'); ax.legend(loc='upper right')
ax.set_title('优化前后关键指标（示意数据）')
save(fig, 'effect')

# 13 训练收敛
fig, ax = plt.subplots(figsize=(5.6, 3.4))
ep = np.arange(1, 101)
ax.plot(ep, 0.9 * np.exp(-ep / 18) + 0.08 + rng.normal(0, 0.008, 100), color=NAVY, lw=2, label='训练集')
ax.plot(ep, 0.95 * np.exp(-ep / 20) + 0.11 + rng.normal(0, 0.012, 100), color=RED, lw=2, label='验证集')
ax.set_xlabel('迭代轮次'); ax.set_ylabel('损失'); ax.legend(); ax.set_title('模型训练收敛（示意数据）')
save(fig, 'loss')

# 14 数据质量：缺失补全前后
fig, ax = plt.subplots(figsize=(6.4, 3.4))
t2 = np.arange(48) / 2
raw = 3 + 1.5 * np.exp(-((t2 - 10) / 3) ** 2) + rng.normal(0, 0.1, 48)
gap = raw.copy(); gap[14:22] = np.nan
ax.plot(t2, raw, color=LB, lw=4, label='补全后')
ax.plot(t2, gap, color=NAVY, lw=2, label='原始（含缺失）')
ax.axvspan(7, 11, color=RED, alpha=0.08)
ax.set_xlabel('时刻 / h'); ax.set_ylabel('负荷 / MW'); ax.legend(loc='upper right')
ax.set_title('缺失数据识别与补全（示意数据）')
save(fig, 'clean')
print('done', len(list(out.glob('*.png'))))
