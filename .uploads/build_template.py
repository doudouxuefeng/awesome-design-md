# -*- coding: utf-8 -*-
"""生成孪生模型回流入模板 Excel"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule

wb = Workbook()

# 样式
HEADER = PatternFill('solid', fgColor='305496')
SUBHEAD = PatternFill('solid', fgColor='D9E1F2')
BASELINE = PatternFill('solid', fgColor='FFF2CC')
INPUT = PatternFill('solid', fgColor='E2EFDA')
WARN = PatternFill('solid', fgColor='F8CBAD')
HDR_FONT = Font(color='FFFFFF', bold=True, size=11)
BOLD = Font(bold=True)
THIN = Side(style='thin', color='BFBFBF')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical='center')
CENTER = Alignment(horizontal='center', vertical='center')

# ============ Sheet1: 整体回流 ============
ws1 = wb.active
ws1.title = '整体回流'
ws1['A1'] = '百度推广孪生模型 - 整体指标回流'
ws1['A1'].font = Font(bold=True, size=14, color='305496')
ws1.merge_cells('A1:G1')
ws1['A2'] = '说明: 黄色=基线(勿改), 绿色=运营填写, 橙色=自动计算; 每周末填一行,共4周'
ws1['A2'].font = Font(italic=True, size=9, color='808080')
ws1.merge_cells('A2:G2')

headers = ['执行周', '实际周消费(元)', '实际线索数(条)', '实际CPL(元)', '在跑计划数', '高CPC词数(>30元)', '执行动作概述']
for i,h in enumerate(headers,1):
    c = ws1.cell(row=4, column=i, value=h)
    c.fill = HEADER; c.font = HDR_FONT; c.alignment = WRAP; c.border = BORDER

# 预填基线
baseline = ['基线(执行前)', 1929.49, 2, 964.75, 16, 12, '孪生模型校准起点']
for i,v in enumerate(baseline,1):
    c = ws1.cell(row=5, column=i, value=v)
    c.fill = BASELINE; c.border = BORDER
    if i>=2 and i<=6: c.alignment = CENTER

# 4 周输入行 (w, cost, lead, cpl, plan, hkw, act) 7-tuple
targets = [
    ('第1周(P1-P5止血)', '~700', '?', '?', '≤9', '≤2', '暂停7计划+否定29词+地域调价'),
    ('第2周(M1-M2提质)', '~1929(含重投1579)', '5-7', '300-400', '9-11', '≤2', '重投优质计划分3批+落地页诊断'),
    ('第3周(M3-M5精修)', '~1929', '6-8', '300-380', '9-11', '≤1', '匹配收紧+清240词+创意A/B'),
    ('第4周(L1拓量)', '~2300', '8-12', '300-350', '10-12', '≤1', '扩词+移动拓量'),
]
for r,(w, cost, lead, cpl, plan, hkw, act) in enumerate(targets, 6):
    ws1.cell(row=r, column=1, value=w).fill = SUBHEAD
    # 输入区
    for col in [2,3,4,5,6]:
        c = ws1.cell(row=r, column=col)
        c.fill = INPUT; c.border = BORDER; c.alignment = CENTER
    # 目标参考(灰色提示)
    ws1.cell(row=r, column=2).comment = None
    # 在消费列填目标参考
    ws1.cell(row=r, column=2, value=cost)
    ws1.cell(row=r, column=3, value=lead)
    ws1.cell(row=r, column=4, value=cpl)
    ws1.cell(row=r, column=5, value=plan)
    ws1.cell(row=r, column=6, value=hkw)
    ws1.cell(row=r, column=7, value=act)
    for col in range(1,8):
        ws1.cell(row=r, column=col).border = BORDER
        ws1.cell(row=r, column=col).alignment = WRAP if col==7 else CENTER

# 备注行
ws1.cell(row=11, column=1, value='填写说明').font = BOLD
notes = [
    '1. 实际周消费 = 后台消费日报 7 日合计',
    '2. 实际线索数 = 综合线索收集总数(留线索+表单+电话)',
    '3. 实际CPL = 实际周消费 / 实际线索数',
    '4. 在跑计划数 = 计划列表中状态=有效的计划数',
    '5. 高CPC词数 = 关键词中有效CPC>30元的数量',
    '6. 执行动作概述 = 本周实际执行的关键动作摘要',
]
for i,n in enumerate(notes,12):
    ws1.cell(row=11+i, column=1, value=n).font = Font(size=9)
    ws1.merge_cells(start_row=11+i, start_column=1, end_row=11+i, end_column=7)

# 列宽
widths = [22, 16, 14, 14, 12, 16, 40]
for i,w in enumerate(widths,1):
    ws1.column_dimensions[get_column_letter(i)].width = w

# ============ Sheet2: 优质计划重投监控 ============
ws2 = wb.create_sheet('重投计划监控')
ws2['A1'] = '优质计划重投分批监控 (M1核心)'
ws2['A1'].font = Font(bold=True, size=14, color='305496')
ws2.merge_cells('A1:G1')
ws2['A2'] = '孪生仿真预警: 单批加码>30% 或 CPL漂移>10% 立即暂停加码'
ws2['A2'].font = Font(italic=True, size=9, color='C00000')
ws2.merge_cells('A2:G2')

h2 = ['批次', '时点', '导入金额(元)', '累计重投(元)', '本周该计划CPL(元)', 'CPL漂移率', '是否继续加码']
for i,h in enumerate(h2,1):
    c = ws2.cell(row=4, column=i, value=h)
    c.fill = HEADER; c.font = HDR_FONT; c.alignment = WRAP; c.border = BORDER

# 2 个优质计划基线
plans = [
    ('2文化pc-ocpc【文化-广播】', 327, 1, 327),
    ('2增值yd-ocpc【消费词】转化', 299, 1, 299),
]
row = 5
for name, base_cost, base_lead, base_cpl in plans:
    ws2.cell(row=row, column=1, value=name).font = BOLD
    ws2.cell(row=row, column=1).fill = SUBHEAD
    ws2.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
    row += 1
    # 基线行
    ws2.cell(row=row, column=1, value='基线(执行前)')
    ws2.cell(row=row, column=2, value='-')
    ws2.cell(row=row, column=3, value=base_cost)
    ws2.cell(row=row, column=4, value=0)
    ws2.cell(row=row, column=5, value=base_cpl)
    ws2.cell(row=row, column=6, value=0)
    ws2.cell(row=row, column=7, value='-')
    for col in range(1,8):
        ws2.cell(row=row, column=col).fill = BASELINE
        ws2.cell(row=row, column=col).border = BORDER
        ws2.cell(row=row, column=col).alignment = CENTER
    row += 1
    # 3 批输入行(参考节奏)
    batches = [
        ('第1批', '第2周D1', 474*0.5, 0.5, '≤400', '≤10%', '观察3天CPL再决定'),
        ('第2批', '第2周D4', 474*0.3, 0.8, '≤420', '≤10%', 'CPL漂移<10%再加'),
        ('第3批', '第2周末', 474*0.2, 1.0, '≤450', '≤15%', '持续监控'),
    ]
    for b,time,amt,acc,cpl_ref,drift_ref,action in batches:
        ws2.cell(row=row, column=1, value=b)
        ws2.cell(row=row, column=2, value=time)
        ws2.cell(row=row, column=3, value=amt)
        ws2.cell(row=row, column=4, value=acc)
        ws2.cell(row=row, column=5, value=cpl_ref)
        ws2.cell(row=row, column=6, value=drift_ref)
        ws2.cell(row=row, column=7, value=action)
        for col in range(1,8):
            ws2.cell(row=row, column=col).fill = INPUT
            ws2.cell(row=row, column=col).border = BORDER
            ws2.cell(row=row, column=col).alignment = CENTER if col<7 else WRAP
        row += 1
    # 空一行
    row += 1

# CPL漂移率公式说明
ws2.cell(row=row, column=1, value='CPL漂移率公式: (本周CPL - 基线CPL) / 基线CPL × 100%').font = Font(size=9, italic=True)
ws2.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
row += 1
ws2.cell(row=row, column=1, value='判断规则: >10% = 警告(暂停加码), >20% = 红线(立即回滚)').font = Font(size=9, italic=True, color='C00000')
ws2.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)

widths2 = [28, 12, 16, 18, 22, 14, 30]
for i,w in enumerate(widths2,1):
    ws2.column_dimensions[get_column_letter(i)].width = w

# ============ Sheet3: 异常事件记录 ============
ws3 = wb.create_sheet('异常事件日志')
ws3['A1'] = '执行期异常事件日志'
ws3['A1'].font = Font(bold=True, size=14, color='305496')
ws3.merge_cells('A1:E1')

h3 = ['日期', '事件类型', '计划/关键词', '现象描述', '处置动作']
for i,h in enumerate(h3,1):
    c = ws3.cell(row=3, column=i, value=h)
    c.fill = HEADER; c.font = HDR_FONT; c.alignment = WRAP; c.border = BORDER

# 示例行
events = [
    ('2026-09-10', '流量断崖', '网文-一二计划暂停后', '点击量-80%', '观察2天,若无线索承接则临时恢复50%预算'),
    ('2026-09-12', 'CPL反弹', '2文化pc重投第2批后', 'CPL从327涨至500+', '暂停第3批加码,排查落地页'),
]
for i,(d,t,o,p,a) in enumerate(events,4):
    ws3.cell(row=i, column=1, value=d)
    ws3.cell(row=i, column=2, value=t)
    ws3.cell(row=i, column=3, value=o)
    ws3.cell(row=i, column=4, value=p)
    ws3.cell(row=i, column=5, value=a)
    for col in range(1,6):
        ws3.cell(row=i, column=col).fill = INPUT
        ws3.cell(row=i, column=col).border = BORDER
        ws3.cell(row=i, column=col).alignment = WRAP

# 10 个空行供运营填
for r in range(6,16):
    for col in range(1,6):
        ws3.cell(row=r, column=col).fill = INPUT
        ws3.cell(row=r, column=col).border = BORDER

widths3 = [12, 14, 28, 36, 36]
for i,w in enumerate(widths3,1):
    ws3.column_dimensions[get_column_letter(i)].width = w

# ============ Sheet4: 校准需求字段 ============
ws4 = wb.create_sheet('孪生校准输入')
ws4['A1'] = '孪生模型重校准 - 必填字段清单'
ws4['A1'].font = Font(bold=True, size=14, color='305496')
ws4.merge_cells('A1:D1')
ws4['A2'] = '运营回流这些字段后, AI 将重新校准 CTR/CPC/CVR 参数, 更新仿真区间'
ws4['A2'].font = Font(italic=True, size=9, color='808080')
ws4.merge_cells('A2:D2')

h4 = ['字段名', '本周值', '获取路径', '用途']
for i,h in enumerate(h4,1):
    c = ws4.cell(row=4, column=i, value=h)
    c.fill = HEADER; c.font = HDR_FONT; c.alignment = WRAP; c.border = BORDER

fields = [
    ('总展现量', '', '后台-数据报告-关键词报告', '校准 CTR'),
    ('总点击量', '', '同上', '校准 CTR + CPC'),
    ('总消费', '', '同上', '校准 CPC + CPL'),
    ('综合线索总数', '', '同上', '校准 CVR + CPL'),
    ('在跑计划数', '', '计划管理', '结构变化监测'),
    ('TOP3计划消费占比%', '', '计划报告', '集中度监测'),
    ('0线索消费占比%', '', '关键词报告筛选', '浪费率监测'),
    ('重投计划(2文化pc)CPL', '', '计划报告', '优质计划漂移监测'),
    ('重投计划(2增值yd)CPL', '', '计划报告', '优质计划漂移监测'),
    ('落地页是否改版', '是/否', '创意管理', 'M2 收益归因'),
]
for i,(f,v,src,u) in enumerate(fields,5):
    ws4.cell(row=i, column=1, value=f).font = BOLD
    c = ws4.cell(row=i, column=2, value=v); c.fill = INPUT; c.alignment = CENTER
    ws4.cell(row=i, column=3, value=src).alignment = WRAP
    ws4.cell(row=i, column=4, value=u).alignment = WRAP
    for col in range(1,5):
        ws4.cell(row=i, column=col).border = BORDER

widths4 = [26, 16, 28, 28]
for i,w in enumerate(widths4,1):
    ws4.column_dimensions[get_column_letter(i)].width = w

# 冻结首行
for ws in [ws1, ws2, ws3, ws4]:
    ws.freeze_panes = 'A5' if ws is ws1 else ('A5' if ws is ws2 else ('A4' if ws is ws3 else 'A5'))

out = '/workspace/孪生模型回流入模板.xlsx'
wb.save(out)
print(f'OK -> {out}')
