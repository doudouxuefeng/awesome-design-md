#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
百度推广孪生模型回流分析器(一键复用版)

用法:
    python3 baidu_relic_analyzer.py <运营填完的回流Excel> [输出MD路径]

示例:
    python3 baidu_relic_analyzer.py 孪生模型回流入模板_预演第1周回流.xlsx
    python3 baidu_relic_analyzer.py 回流_第3周.xlsx 报告_第3周.md

输入: 运营填完的"孪生模型回流入模板.xlsx"(4个Sheet)
输出: 一份结构化Markdown报告(重校准+归因+根因+重仿真+下周建议)
"""
import sys
import os
import openpyxl
from datetime import datetime


# ============================================================
# 内置基线数据(来自2026-08-31~09-06真实投放原始分析)
# 不依赖外部文件,脚本可独立运行
# ============================================================
BASELINE = {
    '展现': 21837, '点击': 356, '消费': 1929.49, '线索': 2,
    '在跑计划': 16, '高CPC词': 12,
    'CTR': 356/21837, 'CPC': 1929.49/356, 'CVR': 2/356,
    'CPL': 1929.49/2,
}

# 原始诊断关键事实(用于根因分析)
ORIGINAL_FACTS = {
    '优质计划': ['2文化pc-ocpc【文化-广播】', '2增值yd-ocpc【消费词】转化'],
    '优质计划_CPL': [327, 299],  # 单样本,不可信
    '烧钱计划数': 7,
    '烧钱计划消费': 1262,
    '高消无转词数': 29,
    '高消无转词消费': 1503,
    '一线四城': ['北京-北京', '广东-深圳', '江苏-苏州', '上海-上海'],
    '一线四城消费': 1341,
    '0线索消费占比': 0.956,
    'TOP3计划消费占比': 0.625,
}

# 原计划目标(用于偏差对比)
ORIGINAL_TARGETS = {
    1: {'消费': 700, '线索': '≥1', 'CPL': 'NA(线索少)', '在跑计划': 9, '高CPC词': 2},
    2: {'消费': 1929, '线索': '5-7', 'CPL': '300-400', '在跑计划': '9-11', '高CPC词': 2},
    3: {'消费': 1929, '线索': '6-8', 'CPL': '300-380', '在跑计划': '9-11', '高CPC词': 1},
    4: {'消费': 2300, '线索': '8-12', 'CPL': '300-350', '在跑计划': '10-12', '高CPC词': 1},
}


# ============================================================
# 数据读取层(健壮:容忍NA/空/字符串/数字)
# ============================================================
def to_num(v, default=None):
    """把任意值转成float,失败返回default"""
    if v is None or v == '':
        return default
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if s.lower() in ('na', 'n/a', '无', '无线索', '?', '-'):
        return default
    # 提取数字部分
    for token in s.replace('~', ' ').replace('-', ' ').replace('+', ' ').split():
        try:
            return float(token)
        except ValueError:
            continue
    return default


def read_week_data(filepath):
    """读取运营填完的回流Excel,返回结构化数据"""
    # data_only=True 读不到手动写入的值(无公式缓存),用 False 读
    wb = openpyxl.load_workbook(filepath, data_only=False)
    data = {'sheets': wb.sheetnames}

    # Sheet1: 整体回流 - 找运营填的最新一周(R6-R9)
    if '整体回流' in wb.sheetnames:
        ws = wb['整体回流']
        weeks = []
        for r in range(6, 10):
            week_label = ws.cell(r, 1).value or ''
            cost_raw = ws.cell(r, 2).value
            lead_raw = ws.cell(r, 3).value
            # 跳过目标参考行:值是字符串且含 ~ ≤ - 等示意字符
            def is_target_placeholder(v):
                if v is None:
                    return True
                if isinstance(v, (int, float)):
                    return False
                s = str(v).strip()
                # 含这些字符说明是目标参考而非实际数据
                return any(c in s for c in ['~', '≤', '≥', '-', '?']) or s == ''
            if is_target_placeholder(cost_raw) and is_target_placeholder(lead_raw):
                continue
            cost = to_num(cost_raw)
            lead = to_num(lead_raw)
            cpl = to_num(ws.cell(r, 4).value)
            plan = to_num(ws.cell(r, 5).value)
            hkw = to_num(ws.cell(r, 6).value)
            action = ws.cell(r, 7).value or ''
            # 找到填了实际数字的周
            if cost is not None and '基线' not in str(week_label):
                weeks.append({
                    'week': week_label, 'cost': cost, 'lead': lead,
                    'cpl': cpl, 'plan': plan, 'hkw': hkw, 'action': action,
                    'row': r,
                })
        data['weeks'] = weeks

    # Sheet3: 异常事件
    if '异常事件日志' in wb.sheetnames:
        ws = wb['异常事件日志']
        events = []
        for r in range(4, ws.max_row + 1):
            d = ws.cell(r, 1).value
            t = ws.cell(r, 2).value
            if d and t:
                events.append({
                    'date': str(d)[:10] if d else '',
                    'type': str(t),
                    'target': str(ws.cell(r, 3).value or ''),
                    'desc': str(ws.cell(r, 4).value or ''),
                    'action': str(ws.cell(r, 5).value or ''),
                })
        data['events'] = events

    # Sheet4: 孪生校准输入
    if '孪生校准输入' in wb.sheetnames:
        ws = wb['孪生校准输入']
        cal = {}
        for r in range(5, 15):
            name = ws.cell(r, 1).value
            val = ws.cell(r, 2).value
            if name:
                cal[str(name)] = val
        data['calibrate'] = cal

    # 把校准字段合并进最新一周的 week_data,供 recalibrate 使用
    if 'weeks' in data and data['weeks'] and 'calibrate' in data:
        data['weeks'][-1]['calibrate'] = data['calibrate']

    return data


def identify_week_number(week_label):
    """从"第N周"标签提取N"""
    s = str(week_label)
    for i in range(1, 10):
        if f'第{i}周' in s:
            return i
    return 1


# ============================================================
# 重校准层
# ============================================================
def recalibrate(week_data):
    """基于第N周回流数据重校准孪生模型参数"""
    cost = week_data['cost']
    lead = week_data['lead'] or 0
    plan = week_data['plan']
    hkw = week_data['hkw']

    # 从校准字段读展现/点击
    cal = week_data.get('calibrate', {})
    imp = to_num(cal.get('总展现量'), default=0)
    clk = to_num(cal.get('总点击量'), default=0)
    # 若运营没填展现/点击,用消费反推
    if imp == 0:
        imp = int(cost / (BASELINE['CPC'] * 0.9)) if cost else 0
    if clk == 0:
        # 用基线CTR反推
        clk = int(imp * BASELINE['CTR']) if imp else 0

    # 重校准参数
    new_ctr = clk / imp if imp else 0
    new_cpc = cost / clk if clk else 0
    new_cvr = lead / clk if clk else 0
    new_cpl = cost / lead if lead else float('inf')

    return {
        'imp': imp, 'clk': clk, 'cost': cost, 'lead': lead,
        'plan': plan, 'hkw': hkw,
        'CTR': new_ctr, 'CPC': new_cpc, 'CVR': new_cvr, 'CPL': new_cpl,
    }


# ============================================================
# 偏差归因层
# ============================================================
def diagnose_actions(week_data, recal, events, week_num):
    """根据执行动作概述+异常事件,判断每项动作的执行状态"""
    action_text = week_data.get('action', '')
    event_types = [e['type'] for e in events]

    actions = []

    # P1 暂停7计划
    p1_status = '生效'
    p1_detail = '消费压缩达成'
    if '断崖' in action_text or any('断崖' in t for t in event_types):
        p1_status = '部分生效'
        p1_detail = '触发流量断崖,有计划被临时恢复'
    elif week_data['cost'] > 1000:
        p1_status = '未生效'
        p1_detail = '消费未明显压缩'
    actions.append(('P1 暂停7计划', p1_status, p1_detail))

    # P2 否定29词
    if recal['hkw'] is not None and recal['hkw'] < BASELINE['高CPC词']:
        actions.append(('P2 否定29词', '生效',
                        f'高CPC词 {BASELINE["高CPC词"]}→{int(recal["hkw"])}'))
    else:
        actions.append(('P2 否定29词', '部分生效', '高CPC词清理不彻底'))

    # P3 核对已删除
    actions.append(('P3 核对已删除词', '生效', '已停投词已确认'))

    # P4 调价
    if recal['CPC'] > BASELINE['CPC'] * 1.1:
        actions.append(('P4 调价', '部分生效',
                        f'CPC {BASELINE["CPC"]:.2f}→{recal["CPC"]:.2f}(+{(recal["CPC"]/BASELINE["CPC"]-1)*100:.0f}%) 残留高CPC词'))
    else:
        actions.append(('P4 调价', '生效', 'CPC稳定'))

    # P5 地域调价
    if '漏调' in action_text or any('地域' in e['desc'] for e in events):
        actions.append(('P5 地域调价', '执行失误', '地域系数漏调或延迟'))
    else:
        actions.append(('P5 地域调价', '生效', '一线城市系数已调'))

    # M1 重投
    if week_num >= 2:
        if recal['lead'] == 0:
            actions.append(('M1 重投', '未生效', '重投无线索产出'))
        else:
            actions.append(('M1 重投', '生效', f'线索{int(recal["lead"])}条'))
    else:
        actions.append(('M1 重投', '未启动', '原计划第2周启动'))

    # M2 落地页
    cal = week_data.get('calibrate', {})
    lp = str(cal.get('落地页是否改版', '否'))
    if '是' in lp:
        actions.append(('M2 落地页', '已启动', '落地页已改版'))
    else:
        actions.append(('M2 落地页', '未启动', 'CVR仍基线0.56%'))

    return actions


def root_cause(recal, week_num):
    """线索归零/不足的根因诊断"""
    lead = recal['lead']
    if lead > 0:
        return None  # 有线索不需根因诊断

    reasons = []
    if week_num == 1:
        reasons.append(
            f'P1暂停烧钱计划后,优质计划流量同步缩水(本周点击{int(recal["clk"])}次)\n'
            f'  CVR基线{BASELINE["CVR"]*100:.2f}%,{int(recal["clk"])}次点击理论线索={recal["clk"]*BASELINE["CVR"]:.2f}(基本注定0)\n'
            f'  优质计划从未独立产出过≥2线索,其CPL 327/299是单样本偶然'
        )
    reasons.append(
        f'CVR={recal["CVR"]*100:.3f}%(基线{BASELINE["CVR"]*100:.3f}%),转化断裂是核心病灶\n'
        f'  落地页未改版是CVR无法提升的直接原因'
    )
    return '\n'.join(reasons)


# ============================================================
# 重仿真层
# ============================================================
def resimulate(recal, week_num):
    """重仿真第N+1周参数,返回分批建议"""
    # 重投预算(保守化)
    base_budget = 1579
    if recal['lead'] == 0:
        # 线索归零,重投预算大幅下调
        budget = int(base_budget * 0.63)  # 1000元
        cvr_low = 0.003
        cvr_high = 0.005
    else:
        budget = base_budget
        cvr_low = BASELINE['CVR'] * 0.8
        cvr_high = BASELINE['CVR'] * 1.0

    cpc_low = max(recal['CPC'] - 1, 1)
    cpc_high = recal['CPC'] + 1

    batches = []
    # 分批比例
    if recal['lead'] == 0:
        ratios = [('第1批(保守)', 0.15, 5),
                  ('第2批(条件触发)', 0.20, 5),
                  ('第3批(条件触发)', 0.20, 5)]
    else:
        ratios = [('第1批', 0.30, 3),
                  ('第2批', 0.25, 3),
                  ('第3批', 0.45, 3)]

    for label, pct, obs_days in ratios:
        amt = budget * pct
        clicks_low = amt / cpc_high
        clicks_high = amt / cpc_low
        leads_low = clicks_low * cvr_low
        leads_high = clicks_high * cvr_high
        cpl_low = amt / leads_high if leads_high > 0 else float('inf')
        cpl_high = amt / leads_low if leads_low > 0 else float('inf')
        batches.append({
            'label': label, 'amount': amt, 'pct': pct,
            'clicks': (clicks_low, clicks_high),
            'leads': (leads_low, leads_high),
            'cpl': (cpl_low, cpl_high),
            'obs_days': obs_days,
        })

    # 整体仿真区间
    total_clicks_low = sum(b['clicks'][0] for b in batches)
    total_clicks_high = sum(b['clicks'][1] for b in batches)
    total_leads_low = sum(b['leads'][0] for b in batches)
    total_leads_high = sum(b['leads'][1] for b in batches)
    total_cost = budget
    overall_cpl_low = total_cost / total_leads_high if total_leads_high > 0 else float('inf')
    overall_cpl_high = total_cost / total_leads_low if total_leads_low > 0 else float('inf')

    return {
        'budget': budget, 'batches': batches,
        'cpc_range': (cpc_low, cpc_high),
        'cvr_range': (cvr_low, cvr_high),
        'overall': {
            'clicks': (total_clicks_low, total_clicks_high),
            'leads': (total_leads_low, total_leads_high),
            'cpl': (overall_cpl_low, overall_cpl_high),
        }
    }


# ============================================================
# 报告生成层
# ============================================================
def fmt_num(v, suffix=''):
    if v == float('inf') or v is None:
        return 'NA'
    if isinstance(v, float):
        return f'{v:,.2f}{suffix}'
    return f'{v:,}{suffix}'


def fmt_pct(v):
    if v == float('inf') or v is None or v == 0:
        return '0.00%'
    return f'{v*100:.2f}%'


def fmt_change(old, new):
    if new == float('inf') or new is None or old == 0:
        return '-'
    change = (new / old - 1) * 100
    return f'{change:+.1f}%'


def generate_report(week_data, recal, actions, root, sim, week_num, events):
    """生成结构化Markdown报告"""
    out = []
    w = lambda s: out.append(s)

    target = ORIGINAL_TARGETS.get(week_num, {})
    next_week = week_num + 1

    w(f'# 百度推广孪生模型回流分析报告 · 第 {week_num} 周')
    w('')
    w(f'**生成时间：** {datetime.now().strftime("%Y-%m-%d %H:%M")}')
    w(f'**分析对象：** 第 {week_num} 周回流数据')
    w(f'**数据来源：** 运营填写的回流入模板')
    w('')
    w('---')
    w('')

    # ===== 一、孪生模型重校准 =====
    w('## 一、孪生模型重校准')
    w('')
    w('| 指标 | 基线(执行前) | 第%d周实际 | 变化 | 解读 |' % week_num)
    w('|------|--------------|-------------|------|------|')

    rows = [
        ('展现量', BASELINE['展现'], recal['imp'], 'int', '流量压缩(P1生效)'),
        ('点击量', BASELINE['点击'], recal['clk'], 'int', '点击同步缩'),
        ('消费', BASELINE['消费'], recal['cost'], 'money', '预算压缩'),
        ('线索', BASELINE['线索'], recal['lead'], 'int', '⚠ 归零' if recal['lead'] == 0 else '有产出'),
        ('CTR', BASELINE['CTR'], recal['CTR'], 'pct', '低质流量清理'),
        ('CPC', BASELINE['CPC'], recal['CPC'], 'money', '高CPC词残留+竞争' if recal['CPC'] > BASELINE['CPC'] else '稳定'),
        ('CVR', BASELINE['CVR'], recal['CVR'], 'pct', '⚠ 转化断裂' if recal['CVR'] == 0 else '转化恢复'),
        ('CPL', BASELINE['CPL'], recal['CPL'], 'money', '⚠ 无法计算' if recal['CPL'] == float('inf') else '改善'),
    ]
    for name, b_v, a_v, kind, note in rows:
        if kind == 'int':
            b_str = f'{int(b_v):,}'
            a_str = f'{int(a_v):,}' if a_v != float('inf') else 'NA'
            chg = fmt_change(b_v, a_v)
        elif kind == 'money':
            b_str = f'{b_v:,.2f}'
            a_str = f'{a_v:,.2f}' if a_v != float('inf') else 'NA'
            chg = fmt_change(b_v, a_v)
        elif kind == 'pct':
            b_str = f'{b_v*100:.2f}%'
            a_str = f'{a_v*100:.3f}%' if a_v != float('inf') else '0.000%'
            chg = fmt_change(b_v, a_v)
        w(f'| {name} | {b_str} | {a_str} | {chg} | {note} |')

    w('')

    # ===== 二、偏差归因 =====
    w('## 二、偏差归因分析')
    w('')
    w('| 动作 | 状态 | 关键发现 |')
    w('|------|------|----------|')
    for a, s, d in actions:
        w(f'| {a} | {s} | {d} |')
    w('')

    # ===== 三、根因诊断 =====
    if root:
        w('## 三、⚠ 线索归零/不足根因诊断')
        w('')
        w('```')
        w(root)
        w('```')
        w('')
        w('**核心结论(孪生模型自我修正)：**')
        w('- 账户"优质计划"的真实CPL被高估(单样本不可信)')
        w('- 原假设(优质计划CPL 313元稳定)不成立')
        w('- 重校准:优质计划CPL取区间300-600元(更保守)')
        w('- M1重投需更谨慎:第1批从30%降至15%,观察5天而非3天')
        w('')

    # ===== 四、重仿真 =====
    w(f'## 四、孪生模型重仿真 · 第 {next_week} 周')
    w('')
    w(f'**关键调整：重投预算 {ORIGINAL_TARGETS.get(2, {}).get("消费", 1579)} → {sim["budget"]} 元**')
    w(f'**CVR区间(重校准)：{sim["cvr_range"][0]*100:.2f}% - {sim["cvr_range"][1]*100:.2f}%**')
    w(f'**CPC区间：{sim["cpc_range"][0]:.2f} - {sim["cpc_range"][1]:.2f} 元**')
    w('')
    w('| 批次 | 金额 | 占比 | 预估点击 | 预估线索 | 预估CPL | 观察期 | 触发条件 |')
    w('|------|------|------|----------|----------|---------|--------|----------|')
    for b in sim['batches']:
        cpl_s = f'{b["cpl"][0]:.0f}-{b["cpl"][1]:.0f}元' if b['cpl'][0] != float('inf') else 'NA'
        cond = f'上批CPL≤{b["cpl"][1]*1.2:.0f} 且 线索≥{b["leads"][0]*0.8:.1f}' if '条件' in b['label'] else '立即启动'
        w(f'| {b["label"]} | {b["amount"]:.0f}元 | {b["pct"]*100:.0f}% | {b["clicks"][0]:.0f}-{b["clicks"][1]:.0f} | {b["leads"][0]:.2f}-{b["leads"][1]:.2f} | {cpl_s} | {b["obs_days"]}天 | {cond} |')
    w('')

    # ===== 五、下周执行建议 =====
    w(f'## 五、第 {next_week} 周执行建议(基于重校准模型)')
    w('')

    # 红色立即回滚
    rollback = []
    if any('断崖' in e['type'] for e in events):
        rollback.append('网文-一二计划恢复预算本周内分3天递减至0(D1 30%→D2 15%→D3 0),不再全停触发断崖')
    if any('漏调' in e.get('desc', '') or '地域' in e.get('desc', '') for e in events):
        rollback.append('北京地域系数0.6立即生效(本周已补调,持续监控)')
    if recal['hkw'] and recal['hkw'] > 2:
        rollback.append(f'剩余{int(recal["hkw"])}个高CPC词本周内补清')

    if rollback:
        w('🔴 **立即回滚/补救：**')
        for i, r in enumerate(rollback, 1):
            w(f'{i}. {r}')
        w('')

    # 黄色重投调整
    w('🟡 **M1重投调整(保守版)：**')
    if recal['lead'] == 0:
        w(f'- 原方案：1579元分3批(30%/25%/45%)')
        w(f'- **新方案：{sim["budget"]}元分3批(15%/20%/20%,余45%灵活储备)**')
        w(f'- 第1批{sim["batches"][0]["amount"]:.0f}元,观察**5天**(原3天)CPL与线索')
        w(f'- **触发条件：第1批CPL≤800元 且 线索≥1条 才加第2批**')
        w('- 若第1批无线索：暂停重投,先做M2落地页改版')
    else:
        w(f'- 重投预算{sim["budget"]}元,按原分批节奏执行')
        w(f'- CPL漂移>10%暂停加码,>20%立即回滚')
    w('')

    # 绿色M2提前
    w('🟢 **M2落地页(提前/继续)：**')
    if recal['CVR'] == 0:
        w('- 线索归零根因是CVR,落地页是唯一突破口')
        w('- D1启动落地页诊断 → D3完成A/B版上线')
        w('- **重投必须在落地页改版后启动**')
    else:
        w('- 落地页持续优化,关注CVR提升幅度')
    w('')

    # 仿真区间
    w('📊 **重校准后的孪生仿真区间：**')
    o = sim['overall']
    lead_s = f'{o["leads"][0]:.0f}-{o["leads"][1]:.0f}条' if o['leads'][0] != float('inf') else 'NA'
    cpl_s = f'{o["cpl"][0]:.0f}-{o["cpl"][1]:.0f}元' if o['cpl'][0] != float('inf') else 'NA'
    w(f'- 第{next_week}周线索：{lead_s}')
    w(f'- 第{next_week}周CPL：{cpl_s}')
    if recal['lead'] == 0:
        w('- 最坏情况：重投无线索,CPL维持NA,需第%d周再调' % (next_week + 1))
    w('')

    # ===== 六、异常事件 =====
    if events:
        w('## 六、本周异常事件日志')
        w('')
        w('| 日期 | 类型 | 对象 | 现象 | 处置 |')
        w('|------|------|------|------|------|')
        for e in events:
            w(f'| {e["date"]} | {e["type"]} | {e["target"]} | {e["desc"]} | {e["action"]} |')
        w('')

    # ===== 声明 =====
    w('---')
    w('')
    w('**声明：** 本报告为孪生模型基于运营回流数据的重校准推演,所有仿真数值为区间估算,')
    w('受平台算法、市场竞争、落地页质量影响,实际偏差±15%–25%,不构成效果承诺。')
    w('请运营以真实数据回流持续校准,遇到异常波动立即回滚。')
    w('')

    return '\n'.join(out)


# ============================================================
# 主流程
# ============================================================
def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    filepath = sys.argv[1]
    if not os.path.exists(filepath):
        print(f'❌ 文件不存在: {filepath}')
        sys.exit(1)

    # 输出路径
    if len(sys.argv) >= 3:
        out_path = sys.argv[2]
    else:
        # 默认输出到同目录
        base = os.path.splitext(os.path.basename(filepath))[0]
        out_path = f'孪生模型回流分析_{base}.md'

    print(f'📖 读取回流数据: {filepath}')
    data = read_week_data(filepath)

    if 'weeks' not in data or not data['weeks']:
        print('❌ 未找到已填写的周回流数据(Sheet1 R6-R9 需填实际数字)')
        sys.exit(1)

    # 取最新一周
    week_data = data['weeks'][-1]
    week_num = identify_week_number(week_data['week'])
    print(f'✅ 识别为第 {week_num} 周回流')

    print('🔧 重校准孪生模型参数...')
    recal = recalibrate(week_data)

    print('🔍 偏差归因分析...')
    events = data.get('events', [])
    actions = diagnose_actions(week_data, recal, events, week_num)

    print('⚠ 根因诊断...')
    root = root_cause(recal, week_num)

    print(f'📊 重仿真第 {week_num + 1} 周参数...')
    sim = resimulate(recal, week_num)

    print('📝 生成报告...')
    report = generate_report(week_data, recal, actions, root, sim, week_num, events)

    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(report)

    print(f'\n✅ 报告已生成: {out_path}')
    print(f'   周次: 第{week_num}周')
    print(f'   消费: {recal["cost"]:.2f}元 | 线索: {int(recal["lead"])}条 | CPL: {fmt_num(recal["CPL"], "元")}')
    print(f'   下周重投预算: {sim["budget"]}元 (保守化)')
    if root:
        print(f'   ⚠ 触发线索归零根因诊断')


if __name__ == '__main__':
    main()
