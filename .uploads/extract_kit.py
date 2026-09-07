# -*- coding: utf-8 -*-
"""抽取执行套件所需精确名单"""
import pandas as pd, numpy as np, json

PC='/workspace/.uploads/2a20c7ba-8d9c-4ab3-b02f-4cea1eab4431_计算机端.xlsx'
MOB='/workspace/.uploads/62b1e5d1-1889-4240-b681-69826c9be436_移动设备端.xlsx'
cols=['日期','账户','计划','单元','关键词','展现量','点击量','消费','点击率','平均点击价格',
      '线索_留线索','线索_表单','线索_电话','综合线索','城市']
df=pd.concat([pd.read_excel(PC,sheet_name=0),pd.read_excel(MOB,sheet_name=0)],ignore_index=True)
df.columns=cols; df['设备']=['PC']*1921+['移动']*6476
for c in ['展现量','点击量','消费','综合线索']: df[c]=pd.to_numeric(df[c],errors='coerce').fillna(0)
df['平均点击价格']=pd.to_numeric(df['平均点击价格'],errors='coerce').fillna(0)
cost=df['消费'].sum()
def eff_cpc(s):
    c=s['消费'].sum(); k=s['点击量'].sum()
    return c/k if k else 0

# === 1. 7 个烧钱低效计划(精确名单) ===
plan=df.groupby('计划').agg(消费=('消费','sum'),展现=('展现量','sum'),点击=('点击量','sum'),
    线索=('综合线索','sum')).reset_index()
plan['CTR']=plan['点击']/plan['展现']*100
plan['CPC']=plan.apply(lambda r: r['消费']/r['点击'] if r['点击'] else 0, axis=1)
burn=plan[(plan['消费']>50)&(plan['线索']==0)].sort_values('消费',ascending=False)
print('==== P1: 烧钱低效计划(暂停) ====')
for _,r in burn.iterrows():
    print(f'{r["计划"]}\t{r["消费"]:.0f}\t{r["展现"]:.0f}\t{r["点击"]:.0f}\t{r["CTR"]:.2f}%\t{r["CPC"]:.2f}')

# === 2. 全部 29 个高消无转关键词(精确名单,用于否定) ===
kw=df.groupby('关键词').agg(消费=('消费','sum'),展现=('展现量','sum'),点击=('点击量','sum'),
    线索=('综合线索','sum')).reset_index()
kw['CTR']=kw['点击']/kw['展现']*100
kw['CPC']=kw.apply(lambda r: r['消费']/r['点击'] if r['点击'] else 0, axis=1)
noconv=kw[(kw['消费']>20)&(kw['线索']==0)].sort_values('消费',ascending=False)
print('\n==== P2: 高消无转关键词(全部29个,用于精确否定/暂停) ====')
for i,(_,r) in enumerate(noconv.iterrows(),1):
    print(f'{i}\t{r["关键词"]}\t{r["消费"]:.0f}\t{r["点击"]:.0f}\t{r["CTR"]:.2f}%\t{r["CPC"]:.2f}')

# === 3. [已删除] 关键词 ===
deleted=kw[kw['关键词'].astype(str).str.contains('已删除', na=False)].sort_values('消费',ascending=False)
print('\n==== P3: [已删除]关键词核对(确认是否真停投) ====')
for _,r in deleted.iterrows():
    print(f'{r["关键词"]}\t{r["消费"]:.0f}\t{r["点击"]:.0f}\t{r["CPC"]:.2f}')

# === 4. PC端高CPC关键词(CPC>30, 出价下调) ===
highcpc=kw[(kw['CPC']>30)&(kw['消费']>20)].sort_values('CPC',ascending=False)
print('\n==== P4: 高CPC关键词(CPC>30元,出价下调50%) ====')
for _,r in highcpc.iterrows():
    print(f'{r["关键词"]}\t{r["消费"]:.0f}\t{r["点击"]:.0f}\t{r["CPC"]:.2f}')

# === 5. 高消费无转化地域 ===
geo=df.groupby('城市').agg(消费=('消费','sum'),展现=('展现量','sum'),
    点击=('点击量','sum'),线索=('综合线索','sum')).reset_index()
geo=geo.sort_values('消费',ascending=False)
geo_burn=geo[(geo['消费']>100)&(geo['线索']==0)]
print('\n==== P5: 高消费无转化地域(出价下调40%) ====')
for _,r in geo_burn.iterrows():
    print(f'{r["城市"]}\t{r["消费"]:.0f}\t{r["展现"]:.0f}\t{r["线索"]:.0f}')

# === 6. 240个0点击有展现关键词 TOP20(按展现) ===
zero_click=kw[(kw['展现']>=5)&(kw['点击']==0)].sort_values('展现',ascending=False)
print(f'\n==== M4: 0点击有展现关键词(共{len(zero_click)}个,展现TOP20清理) ====')
for i,(_,r) in enumerate(zero_click.head(20).iterrows(),1):
    print(f'{i}\t{r["关键词"]}\t{r["展现"]:.0f}')

# === 7. 优质计划重投目标明细 ===
good=plan[plan['线索']>0].copy()
good['CPL']=good['消费']/good['线索']
good=good.sort_values('消费',ascending=False)
print('\n==== M1: 优质计划重投目标 ====')
for _,r in good.iterrows():
    print(f'{r["计划"]}\t{r["消费"]:.0f}\t{r["线索"]:.0f}\t{r["CPL"]:.1f}')

# === 8. 优质计划内的关键词明细(用于扩词/加价参考) ===
for p in good['计划']:
    sub=df[df['计划']==p].groupby('关键词').agg(消费=('消费','sum'),点击=('点击量','sum'),
        线索=('综合线索','sum'),CPC=('平均点击价格','mean')).reset_index()
    sub=sub[sub['消费']>0].sort_values('消费',ascending=False)
    print(f'\n==== 优质计划[{p}]内关键词明细 ====')
    for _,r in sub.head(8).iterrows():
        print(f'{r["关键词"]}\t{r["消费"]:.0f}\t{r["点击"]:.0f}\t{r["线索"]:.0f}\t{r["CPC"]:.2f}')

# === 9. 产出线索的唯一关键词(培育起点) ===
lead_kw=kw[kw['线索']>0].sort_values('线索',ascending=False)
print('\n==== 培育起点: 已产出线索关键词 ====')
for _,r in lead_kw.iterrows():
    print(f'{r["关键词"]}\t{r["消费"]:.0f}\t{r["点击"]:.0f}\t{r["线索"]:.0f}\t{r["消费"]/r["线索"]:.1f}')

# 落盘 CSV 否定词清单
noconv[['关键词','消费','点击','CTR','CPC']].to_csv('/workspace/否定关键词清单.csv',
    index=False, encoding='utf-8-sig', float_format='%.2f')
burn[['计划','消费','展现','点击','CTR','CPC']].to_csv('/workspace/暂停计划清单.csv',
    index=False, encoding='utf-8-sig', float_format='%.2f')
print('\nCSV已生成: 否定关键词清单.csv / 暂停计划清单.csv')
