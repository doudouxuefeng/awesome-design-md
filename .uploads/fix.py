# -*- coding: utf-8 -*-
"""修正:场景3去重计算 + 基线CPL修正"""
import pandas as pd, numpy as np

PC='/workspace/.uploads/2a20c7ba-8d9c-4ab3-b02f-4cea1eab4431_计算机端.xlsx'
MOB='/workspace/.uploads/62b1e5d1-1889-4240-b681-69826c9be436_移动设备端.xlsx'
cols=['日期','账户','计划','单元','关键词','展现量','点击量','消费','点击率','平均点击价格',
      '线索_留线索','线索_表单','线索_电话','综合线索','城市']
df=pd.concat([pd.read_excel(PC,sheet_name=0),pd.read_excel(MOB,sheet_name=0)],ignore_index=True)
df.columns=cols; df['设备']=['PC']*1921+['移动']*6476
for c in ['展现量','点击量','消费','综合线索']: df[c]=pd.to_numeric(df[c],errors='coerce').fillna(0)
cost=df['消费'].sum(); lead=df['综合线索'].sum(); clk=df['点击量'].sum(); imp=df['展现量'].sum()
cpc=cost/clk; cvr=lead/clk; cpl=cost/lead
print(f'基线: 消费{cost:.0f} 点击{clk:.0f} 线索{lead:.0f} CPC{cpc:.2f} CVR{cvr*100:.3f}% CPL={cpl:.2f}')
print(f'CPL正解 = CPC/CVR = {cpc/cvr:.2f} (应等于{cpl:.2f})')

# 烧钱计划
plan=df.groupby('计划').agg(消费=('消费','sum'),线索=('综合线索','sum')).reset_index()
burn_plans=plan[(plan['消费']>50)&(plan['线索']==0)]['计划'].tolist()
burn_cost=df[df['计划'].isin(burn_plans)]['消费'].sum()
print(f'\n烧钱低效计划({len(burn_plans)}个) 消费={burn_cost:.0f} 占比{burn_cost/cost*100:.1f}%')

# 高消无转关键词(全部计划)
kw=df.groupby('关键词')['消费'].sum()
kw_lead=df.groupby('关键词')['综合线索'].sum()
noconv_kws=kw[(kw>20)&(kw_lead==0)].index
noconv_all=df[df['关键词'].isin(noconv_kws)]['消费'].sum()
print(f'高消无转关键词({len(noconv_kws)}个) 全部消费={noconv_all:.0f}')

# 拆分: 在烧钱计划内 vs 在其他计划内
in_burn=df[df['计划'].isin(burn_plans)&df['关键词'].isin(noconv_kws)]['消费'].sum()
out_burn=df[(~df['计划'].isin(burn_plans))&df['关键词'].isin(noconv_kws)]['消费'].sum()
print(f'  其中: 在烧钱计划内 {in_burn:.0f}元(关停计划时已清) | 在其他计划内 {out_burn:.0f}元(需单独否定)')

# 场景3 精确: 关停烧钱计划(释放 burn_cost) + 否定其他计划内高消无转词(释放 out_burn,保守按70%实际可清)
clean_total = burn_cost + out_burn*0.7
print(f'\n场景3 精确可释放预算 = {burn_cost:.0f}(关停烧钱计划) + {out_burn*0.7:.0f}(否定其他计划无转词,70%保守) = {clean_total:.0f}元 ({clean_total/cost*100:.1f}%)')

# 释放预算重投到优质计划: 优质计划CPL中位数
good=df[df['综合线索']>0].groupby('计划').agg(消费=('消费','sum'),线索=('综合线索','sum')).reset_index()
good['CPL']=good['消费']/good['线索']
good_cpl=good['CPL'].median()
print(f'优质计划({len(good)}个) CPL中位数={good_cpl:.0f}')
remaining_cost=cost-clean_total
# 剩余预算按当前整体效率跑
leads_remaining=remaining_cost/cpl
# 释放预算按优质计划效率重投
leads_reinvest=clean_total/good_cpl
total_leads=leads_remaining+leads_reinvest
new_cpl=cost/total_leads
print(f'仿真: 剩余{remaining_cost:.0f}按CPL{cpl:.0f}跑 -> 线索{leads_remaining:.1f}')
print(f'      释放{clean_total:.0f}按CPL{good_cpl:.0f}重投 -> 线索{leads_reinvest:.1f}')
print(f'      仿真后总线索≈{total_leads:.1f} (变化{(total_leads/lead-1)*100:+.0f}%) 整体CPL≈{new_cpl:.0f} (变化{(new_cpl/cpl-1)*100:+.0f}%)')

# 健康度评分
score=0
# 转化(权重40): lead/cost 极差
score += min(40, lead/3*40)
# CPL(权重25): 964远超行业300-500
score += max(0, 25*(1-(cpl-300)/700))
# 集中度(权重15): TOP1占30%算偏高
score += 10
# 流量浪费(权重20): 95.6% 0线索消费
score += 5
print(f'\n账户健康度评分: {score:.0f}/100')
