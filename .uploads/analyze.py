# -*- coding: utf-8 -*-
"""百度推广数字孪生分析 - 计算所有指标并落盘"""
import pandas as pd
import numpy as np

PC = '/workspace/.uploads/2a20c7ba-8d9c-4ab3-b02f-4cea1eab4431_计算机端.xlsx'
MOB = '/workspace/.uploads/62b1e5d1-1889-4240-b681-69826c9be436_移动设备端.xlsx'

cols = ['日期','账户','计划','单元','关键词','展现量','点击量','消费','点击率','平均点击价格',
        '线索_留线索','线索_表单','线索_电话','综合线索','城市']

df_pc = pd.read_excel(PC, sheet_name=0)
df_pc.columns = cols
df_pc['设备'] = 'PC'

df_mob = pd.read_excel(MOB, sheet_name=0)
df_mob.columns = cols
df_mob['设备'] = '移动'

df = pd.concat([df_pc, df_mob], ignore_index=True)

# 数值清洗
for c in ['展现量','点击量','消费','线索_留线索','线索_表单','线索_电话','综合线索']:
    df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)

df['日期'] = pd.to_datetime(df['日期'])

out = []
def p(*a):
    out.append(' '.join(str(x) for x in a))

# ============ 1. 基线指标 ============
p('='*60); p('1. 整体基线指标 (PC + 移动 合并)'); p('='*60)
imp = df['展现量'].sum(); clk = df['点击量'].sum(); cost = df['消费'].sum()
lead = df['综合线索'].sum(); lf = df['线索_留线索'].sum()
ls_form = df['线索_表单'].sum(); ls_phone = df['线索_电话'].sum()
ctr = clk/imp*100 if imp else 0
cpc = cost/clk if clk else 0
cvr = lead/clk*100 if clk else 0
cpl = cost/lead if lead else 0
days = df['日期'].nunique()
p(f'数据周期: {df["日期"].min().date()} 至 {df["日期"].max().date()} | {days} 天')
p(f'数据行数: PC={len(df_pc)} 移动={len(df_mob)} 合计={len(df)}')
p(f'账户名称: {df["账户"].iloc[0]}')
p(f'总展现量: {imp:,.0f}')
p(f'总点击量: {clk:,.0f}')
p(f'总消费(元): {cost:,.2f}')
p(f'综合线索总数: {lead:,.0f} (留线索{lf:.0f} / 表单{ls_form:.0f} / 电话{ls_phone:.0f})')
p(f'整体CTR: {ctr:.2f}%')
p(f'平均CPC: {cpc:.2f} 元')
p(f'整体CVR(点击->线索): {cvr:.2f}%')
p(f'平均线索成本(CPL): {cpl:.2f} 元')
p(f'日均消费: {cost/days:,.2f} 元/天')
p(f'日均线索: {lead/days:.2f} 条/天')
# ROI估算(假设客单价区间,服务行业资质办理客单价通常2000-5000)
for v in [1500, 2500, 3500]:
    rev = lead*v
    p(f'  若客单价{v}元 -> 预估营收{rev:,.0f}, ROI(营收/消费)={rev/cost:.2f}')

# ============ 1b. 设备拆分 ============
p(); p('--- 设备拆分 ---')
for dev, g in df.groupby('设备'):
    i,c,co,l = g['展现量'].sum(), g['点击量'].sum(), g['消费'].sum(), g['综合线索'].sum()
    p(f'{dev}: 展现{i:,.0f} 点击{c:,.0f} 消费{co:,.2f} 线索{l:.0f} | CTR{c/i*100:.2f}% CPC{co/c:.2f} CPL{co/l if l else 0:.2f} 消费占比{co/cost*100:.1f}% 线索占比{l/lead*100:.1f}%')

# ============ 2. 计划维度 ============
p(); p('='*60); p('2. 计划维度诊断'); p('='*60)
plan = df.groupby('计划').agg(
    消费=('消费','sum'), 展现=('展现量','sum'), 点击=('点击量','sum'),
    线索=('综合线索','sum'), 行数=('关键词','count')).reset_index()
plan['CTR'] = plan['点击']/plan['展现']*100
plan['CPC'] = plan['消费']/plan['点击']
plan['CPL'] = plan['消费']/plan['线索'].replace(0,np.nan)
plan['消费占比'] = plan['消费']/cost*100
plan['线索占比'] = plan['线索']/lead*100
plan = plan.sort_values('消费', ascending=False)
p(f'计划总数: {len(plan)}')
p('--- TOP10 消费计划 ---')
for _,r in plan.head(10).iterrows():
    cpl_s = f'{r["CPL"]:.1f}' if pd.notna(r['CPL']) else 'NA(无线索)'
    p(f'  {r["计划"][:30]:32s} | 消费{r["消费"]:>8.0f}({r["消费占比"]:.1f}%) 线索{r["线索"]:>5.0f}({r["线索占比"]:.1f}%) CTR{r["CTR"]:.2f}% CPC{r["CPC"]:.2f} CPL={cpl_s}')

# 烧钱低效(消费>0 线索=0)
burn = plan[(plan['消费']>50) & (plan['线索']==0)].sort_values('消费',ascending=False)
p(); p(f'--- 烧钱低效计划(消费>50且无线索, 共{len(burn)}个) ---')
for _,r in burn.head(8).iterrows():
    p(f'  {r["计划"][:30]:32s} | 消费{r["消费"]:>8.0f} 展现{r["展现"]:>6.0f} 点击{r["点击"]:>5.0f} CTR{r["CTR"]:.2f}% 0线索')

# 优质计划(线索>0 且 CPL较低)
good = plan[plan['线索']>0].copy()
good['CPL'] = good['消费']/good['线索']
good = good.sort_values('线索', ascending=False)
p(); p(f'--- 优质计划(有线索, 共{len(good)}个) ---')
for _,r in good.head(8).iterrows():
    p(f'  {r["计划"][:30]:32s} | 消费{r["消费"]:>8.0f} 线索{r["线索"]:>5.0f} CPL{r["CPL"]:.1f} 线索占比{r["线索占比"]:.1f}%')

# 潜力计划(有展现但消费/线索偏低,可拓)
p(); p('--- 潜力计划(CPC低且有展现,可拓量候选) ---')
pot = plan[(plan['展现']>50)&(plan['CPC']<plan['CPC'].median())].sort_values('展现',ascending=False)
for _,r in pot.head(5).iterrows():
    p(f'  {r["计划"][:30]:32s} | 展现{r["展现"]:>6.0f} 点击{r["点击"]:>5.0f} CPC{r["CPC"]:.2f} 线索{r["线索"]:.0f}')

# ============ 3. 关键词维度 ============
p(); p('='*60); p('3. 关键词维度诊断'); p('='*60)
kw = df.groupby('关键词').agg(消费=('消费','sum'), 展现=('展现量','sum'),
    点击=('点击量','sum'), 线索=('综合线索','sum'), 城市数=('城市','nunique')).reset_index()
kw['CTR'] = kw['点击']/kw['展现']*100
kw['CPC'] = kw['消费']/kw['点击']
kw['CPL'] = kw['消费']/kw['线索'].replace(0,np.nan)
kw = kw.sort_values('消费', ascending=False)
p(f'关键词总数: {len(kw)} (含同词不同城市/计划重复行已聚合)')
p(f'有消费关键词数: {(kw["消费"]>0).sum()}')
p(f'零消费有展现关键词数: {((kw["消费"]==0)&(kw["展现"]>0)).sum()}')

p(); p('--- TOP15 高消费关键词 ---')
for _,r in kw.head(15).iterrows():
    cpl_s = f'{r["CPL"]:.1f}' if pd.notna(r['CPL']) else '无转'
    p(f'  {r["关键词"][:22]:24s} | 消费{r["消费"]:>7.0f} 点击{r["点击"]:>5.0f} 线索{r["线索"]:>4.0f} CTR{r["CTR"]:.2f}% CPC{r["CPC"]:.2f} CPL={cpl_s}')

p(); p('--- 高消费无转化关键词(消费>20且线索=0) ---')
noconv = kw[(kw['消费']>20)&(kw['线索']==0)].sort_values('消费',ascending=False)
p(f'数量: {len(noconv)} | 合计浪费消费: {noconv["消费"].sum():.0f} 元 (占总消费{noconv["消费"].sum()/cost*100:.1f}%)')
for _,r in noconv.head(12).iterrows():
    p(f'  {r["关键词"][:22]:24s} | 消费{r["消费"]:>7.0f} 点击{r["点击"]:>5.0f} CTR{r["CTR"]:.2f}% 0线索')

p(); p('--- 高转化好词(线索>=2,按线索数排序) ---')
conv = kw[kw['线索']>=2].sort_values('线索',ascending=False)
p(f'数量: {len(conv)}')
for _,r in conv.head(10).iterrows():
    p(f'  {r["关键词"][:22]:24s} | 线索{r["线索"]:>4.0f} 消费{r["消费"]:>7.0f} CPL{r["CPL"]:.1f} CTR{r["CTR"]:.2f}%')

p(); p('--- 无效词特征(展现多但0点击,占位浪费) ---')
inv = kw[(kw['展现']>=5)&(kw['点击']==0)]
p(f'0点击有展现关键词数: {len(inv)} | 合计展现{inv["展现"].sum():.0f}(占总展现{inv["展现"].sum()/imp*100:.1f}%)')

# 搜索词问题:关键词字面看是否有明显不相关词
p(); p('--- 关键词长度分布(查短词/泛词风险) ---')
kw['词长'] = kw['关键词'].str.len()
for lo,hi in [(1,4),(5,8),(9,12),(13,99)]:
    g = kw[(kw['词长']>=lo)&(kw['词长']<=hi)&(kw['消费']>0)]
    p(f'  词长{lo}-{hi}: {len(g)}词, 消费{g["消费"].sum():.0f}({g["消费"].sum()/cost*100:.1f}%)')

# ============ 4. 地域维度 ============
p(); p('='*60); p('4. 地域维度诊断'); p('='*60)
geo = df.groupby('城市').agg(消费=('消费','sum'),展现=('展现量','sum'),
    点击=('点击量','sum'),线索=('综合线索','sum')).reset_index()
geo['CPL'] = geo['消费']/geo['线索'].replace(0,np.nan)
geo['消费占比'] = geo['消费']/cost*100
geo = geo.sort_values('消费',ascending=False)
p(f'地域(省-市)数: {len(geo)}')
p('--- TOP12 消费地域 ---')
for _,r in geo.head(12).iterrows():
    cpl_s = f'{r["CPL"]:.1f}' if pd.notna(r['CPL']) else '无转'
    p(f'  {r["城市"][:14]:16s} | 消费{r["消费"]:>7.0f}({r["消费占比"]:.1f}%) 线索{r["线索"]:>4.0f} CPL={cpl_s}')

# 高消费无转化地域
g_burn = geo[(geo['消费']>100)&(geo['线索']==0)].sort_values('消费',ascending=False)
p(); p(f'--- 高消费无转化地域(消费>100且线索=0, 共{len(g_burn)}个) ---')
for _,r in g_burn.head(6).iterrows():
    p(f'  {r["城市"][:14]:16s} | 消费{r["消费"]:>7.0f} 展现{r["展现"]:>6.0f} 0线索')

# ============ 4b. 日期趋势 ============
p(); p('--- 每日趋势 ---')
daily = df.groupby('日期').agg(消费=('消费','sum'),点击=('点击量','sum'),线索=('综合线索','sum')).reset_index()
for _,r in daily.iterrows():
    p(f'  {r["日期"].date()} | 消费{r["消费"]:>8.0f} 点击{r["点击"]:>6.0f} 线索{r["线索"]:>4.0f} 日CPL{r["消费"]/r["线索"] if r["线索"] else 0:.1f}')

# ============ 5. 数字孪生模型参数校准 ============
p(); p('='*60); p('5. 数字孪生模型参数校准'); p('='*60)
# 模型: 展现->点击(CTR) -> 消费(CPC) -> 线索(CVR=CPL倒数)
# 用真实数据校准全局参数,再按"计划级弹性"做仿真
eff_kws = df.groupby('关键词').filter(lambda x: x['消费'].sum()>0)
m_ctr = eff_kws['点击量'].sum()/eff_kws['展现量'].sum()
m_cpc = eff_kws['消费'].sum()/eff_kws['点击量'].sum()
m_cvr = eff_kws['综合线索'].sum()/eff_kws['点击量'].sum()
p(f'模型校准(有消费样本): CTR={m_ctr*100:.2f}% CPC={m_cpc:.2f} CVR={m_cvr*100:.2f}%')
p(f'线索成本基线 CPL = 1/(CTR*CVR) * CPC = {1/(m_ctr*m_cvr)*m_cpc:.2f} 元')

# ============ 6. 仿真场景 ============
p(); p('='*60); p('6. 数字孪生仿真场景'); p('='*60)

# 场景1: 预算调整. 设当前日均预算=B0, 仿真 ±20%/-15%
# 假设边际拓量效率递减(d=0.85): 多花20%仅多获17%点击/线索; 减量同理弹性<1
B0 = cost/days
L0 = lead
d_cpc_elastic = 0.85  # 出价弹性: 消费每变1%, 点击变0.85%
p(f'当前日均消费 B0 = {B0:.0f} 元 | 日均线索 L0 = {L0:.1f}')
for label, delta in [('预算+20%', 0.20), ('预算-15%', -0.15)]:
    new_cost = cost*(1+delta)
    # 点击量弹性: 消费增加 a% -> 点击约 a*0.85% (出价上推CPC也涨,实际点击增长<消费增长)
    click_delta = delta * d_cpc_elastic
    new_clicks = clk*(1+click_delta)
    # CPC上移: 消费涨时CPC涨约 (1-click_delta_pct)/(1+消费delta) - 1 ...简化用 0.6 弹性
    new_cpc = new_cost/new_clicks
    # 线索 = 点击 * CVR (假设CVR在拓量时略降5%因泛流量,减量时略升3%因更精准)
    cvr_adj = 0.95 if delta>0 else 1.03
    new_leads = new_clicks * m_cvr * cvr_adj
    new_cpl = new_cost/new_leads
    p(f'  [{label}] 消费{new_cost:,.0f} 点击{new_clicks:,.0f}({click_delta*100:+.1f}%) 线索{new_leads:.0f}({(new_leads/lead-1)*100:+.1f}%) CPC{new_cpc:.2f} CPL{new_cpl:.1f}')

# 场景2: 出价策略 - 下调部分单元出价(降CPC 10%), 假设排名略降展现-8%, 点击-12%, 消费-20%
p(); p('--- 场景2: 出价下调仿真(CPC-10%) ---')
cpc_new = m_cpc*0.9
imp_new = imp*0.92
clk_new = imp_new * m_ctr * 0.96  # 排名降CTR略降
cost_new = clk_new*cpc_new
lead_new = clk_new*m_cvr*1.02  # 更精准CVR略升
p(f'  CPC{cpc_new:.2f} 展现{imp_new:,.0f}(-8%) 点击{clk_new:,.0f}(-~12%) 消费{cost_new:,.0f}({cost_new/cost*100-100:+.1f}%) 线索{lead_new:.0f} CPL{cost_new/lead_new:.1f}')

# 场景3: 清理低效 - 关停 烧钱低效计划 + 否定 高消无转关键词
burn_cost_plan = burn['消费'].sum()
burn_cost_kw = noconv['消费'].sum()
total_clean = burn_cost_plan + burn_cost_kw*0.5  # 去重保守估
p(); p(f'--- 场景3: 清理低效流量仿真 ---')
p(f'  烧钱低效计划消费: {burn_cost_plan:.0f} 元')
p(f'  高消无转关键词消费: {burn_cost_kw:.0f} 元 (去重后预计可清{burn_cost_kw*0.5:.0f})')
p(f'  合计可释放预算: {total_clean:.0f} 元 ({total_clean/cost*100:.1f}%)')
# 释放预算重新分配到优质计划, 优质计划CPL=基线CPL*0.6(更优)
good_cpl = good['CPL'].median() if len(good) else cpl
real_cost = cost - total_clean
real_leads = (cost - total_clean)/cpl  # 剩余预算按原效率
plus_leads = total_clean / good_cpl   # 释放预算按优质效率
total_leads = real_leads + plus_leads
new_cpl3 = cost/total_leads
p(f'  释放预算按优质计划CPL({good_cpl:.0f})重投: 新增线索{plus_leads:.0f}')
p(f'  仿真后: 总线索≈{total_leads:.0f}({(total_leads/lead-1)*100:+.1f}%) 整体CPL≈{new_cpl3:.1f}({(new_cpl3/cpl-1)*100:+.1f}%)')

# 场景4: 增量拓量 - 拓展优质词(假设拓量20%展现, 转化效率基线*0.8)
p(); p('--- 场景4: 增量拓量仿真(优质词扩展现现+20%) ---')
ext_imp = imp*0.20
ext_clicks = ext_imp * m_ctr * 1.1  # 优质词CTR高
ext_cost = ext_clicks * m_cpc * 1.05
ext_leads = ext_clicks * m_cvr * 0.85  # 拓量边际CVR略降
p(f'  新增展现{ext_imp:,.0f} 新增点击{ext_clicks:,.0f} 新增消费{ext_cost:,.0f}')
p(f'  新增线索{ext_leads:.0f} 新增CPL{ext_cost/ext_leads:.1f}')
p(f'  仿真后总: 消费{cost+ext_cost:,.0f}({(ext_cost/cost)*100:+.1f}%) 线索{lead+ext_leads:.0f}({(ext_leads/lead)*100:+.1f}%)')

# ============ 7. 风险指标 ============
p(); p('='*60); p('7. 风险指标(集中度)'); p('='*60)
top1_cost = plan.iloc[0]['消费']
top3_cost = plan.head(3)['消费'].sum()
top5_kw_cost = kw.head(5)['消费'].sum()
p(f'消费TOP1计划占比: {top1_cost/cost*100:.1f}%')
p(f'消费TOP3计划占比: {top3_cost/cost*100:.1f}%')
p(f'消费TOP5关键词占比: {top5_kw_cost/cost*100:.1f}%')
p(f'0线索消费合计: {df[df["综合线索"]==0]["消费"].sum():.0f} 元 (占比{df[df["综合线索"]==0]["消费"].sum()/cost*100:.1f}%)')

# 落盘
with open('/workspace/.uploads/analysis.txt','w',encoding='utf-8') as f:
    f.write('\n'.join(out))
print('DONE -> /workspace/.uploads/analysis.txt')
print(f'lines: {len(out)}')
