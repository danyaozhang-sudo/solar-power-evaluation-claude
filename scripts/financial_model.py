"""
光伏项目投资评估财务模型
========================
使用方法：修改下方"项目参数"区域后直接运行

依赖：numpy
安装：pip install numpy

输出：投资边界 + 出售边界（共2个任务），含每年净利润表和净现金流量表
"""

import numpy as np

# ============================================================
# >>> 项目参数（在此修改） <<<
# ============================================================
capacity_mw = 100              # 装机容量 MW（直流侧）— 山东德州集中式光伏
annual_hours = 1300            # 等效利用小时数（山东德州 GHI~1550kWh/m² × PR~0.80，取低值）
curtailment_rate = 0.03        # 限电率（山东属消纳友好省，近负荷中心，弃光率低）
degradation_rate = 0.005       # 组件年衰减率（0.5%/年）

# 电价（情况一：不参与机制电价竞价 或 情况二：参与机制电价竞价）
# 修改 mlt_price / mlt_ratio 即可切换场景
mlt_price = 0.3800             # 电量加权均价 元/kWh（情况一=中长协/现货加权，情况二=机制电价）
mlt_ratio = 1.0                # 情况一=0.70, 情况二=1.00
spot_price = 0.3000            # 现货交易均价 元/kWh（情况二下忽略）
spot_ratio = 0.0               # 情况一=0.30, 情况二=0.00
green_price = 0.0300           # 绿电溢价 元/kWh（与中长协/现货叠加）
penalty = 0.015                # 两个细则考核 元/kWh
market_fee_rate = 0.02         # 市场费用分摊

# 融资参数
interest_rate = 0.04           # 融资利率
loan_years = 15                # 融资期限
operating_years = 20           # 经营期
residual_rate = 0.05           # 残值率（光伏组件回收价值高于风电）

# 税务参数
vat_rate = 0.13
om_vat_rate = 0.06             # 运维费增值税率（6% 服务为主 / 13% 备件为主）
additional_tax_rate = 0.12
income_tax_rate = 0.25

# 收益率参数
wacc = 0.055                   # 全投资WACC
equity_cost = 0.07             # 资本金成本

# 迭代搜索范围
search_lo = 2.0                # 最低搜索范围 元/W（集中式光伏下限）
search_hi = 5.5                # 最高搜索范围 元/W（含分布式上限）
# ============================================================

capacity_kw = capacity_mw * 1000
capacity_w = capacity_mw * 1_000_000
base_annual_gen_kwh = capacity_kw * annual_hours * (1 - curtailment_rate)

weighted_price = mlt_ratio * mlt_price + spot_ratio * spot_price
effective_price_base = (weighted_price + green_price - penalty) * (1 - market_fee_rate)
vat_revenue_per_kwh = weighted_price + green_price


def calculate(unit_inv, leverage):
    """计算给定单瓦投资和杠杆率的财务指标，返回汇总指标和年度明细"""
    ti = unit_inv * capacity_w
    debt = ti * leverage
    equity = ti - debt
    annual_dep = ti * (1 - residual_rate) / operating_years
    annual_principal_amt = debt / loan_years

    rev = np.zeros(operating_years)
    gen = np.zeros(operating_years)
    om = np.zeros(operating_years)
    ins = np.zeros(operating_years)
    ebitda = np.zeros(operating_years)
    interest = np.zeros(operating_years)
    principal = np.zeros(operating_years)
    nbv_arr = np.zeros(operating_years)

    vat_credit_cf = 0.0
    vat_credit_cf_unlevered = 0.0
    vat_payable = np.zeros(operating_years)
    add_tax_unlevered = np.zeros(operating_years)

    # 逆变器替换残值累加至折旧池（简化处理：作为追加投资在第10年全额折旧）
    inverter_replacement_cost = 0.0

    for i in range(operating_years):
        year = i + 1

        # 逐年衰减后的发电量
        degradation_factor = (1 - degradation_rate) ** i
        gen[i] = base_annual_gen_kwh * degradation_factor
        rev[i] = gen[i] * effective_price_base

        # 运维费分段 + 第10年逆变器替换
        if year <= 5:
            om_rate = 0.02
        elif year <= 10:
            om_rate = 0.025
        else:
            om_rate = 0.03
        om[i] = capacity_w * om_rate

        # 逆变器替换（第10年）
        if year == 10:
            inverter_replacement_cost = capacity_w * 0.05
            om[i] += inverter_replacement_cost

        acc_dep = annual_dep * year
        nbv = max(0, ti - acc_dep)
        nbv_arr[i] = nbv
        ins[i] = nbv * 0.002

        if year <= loan_years:
            rem = max(0, debt - annual_principal_amt * (year - 1))
            interest[i] = rem * interest_rate
            principal[i] = min(annual_principal_amt, rem)

        ebitda[i] = rev[i] - om[i] - ins[i]

        # VAT（有杠杆）—— 融资租赁进项税抵扣保留
        vat_out = gen[i] * vat_revenue_per_kwh * vat_rate / (1 + vat_rate)
        vat_credit = (principal[i] + interest[i]) * vat_rate + om[i] * om_vat_rate
        avail = vat_credit_cf + vat_credit
        vat_payable[i] = max(0, vat_out - avail)
        vat_credit_cf = max(0, avail - vat_out)

        # VAT（无杠杆）—— 不含利息进项税，用于全投资FCF
        vat_credit_unlevered = principal[i] * vat_rate + om[i] * om_vat_rate
        avail_unlevered = vat_credit_cf_unlevered + vat_credit_unlevered
        vat_payable_unlevered = max(0, vat_out - avail_unlevered)
        vat_credit_cf_unlevered = max(0, avail_unlevered - vat_out)
        add_tax_unlevered[i] = vat_payable_unlevered * additional_tax_rate

    ebit = ebitda - annual_dep
    add_tax = vat_payable * additional_tax_rate
    ebt = ebit - interest - add_tax
    income_tax = np.array([max(0, e) * income_tax_rate for e in ebt])
    net_profit = ebt - income_tax

    # DSCR
    dscr = np.zeros(operating_years)
    for i in range(operating_years):
        total_ds = principal[i] + interest[i]
        if total_ds > 0:
            avail_ds = ebitda[i] - income_tax[i] - add_tax[i]
            dscr[i] = avail_ds / total_ds
        else:
            dscr[i] = float('inf')

    # 净现金流量
    net_cf = ebitda - add_tax - income_tax - interest - principal

    # 全投资FCF（使用无杠杆附加税）
    fcf = np.zeros(operating_years + 1)
    fcf[0] = -ti
    for i in range(operating_years):
        t = max(0, (ebit[i] - add_tax_unlevered[i]) * income_tax_rate)
        fcf[i + 1] = ebitda[i] - add_tax_unlevered[i] - t

    # 资本金FCF
    eq_cf = np.zeros(operating_years + 1)
    eq_cf[0] = -equity
    for i in range(operating_years):
        if leverage < 1.0:
            eq_cf[i + 1] = net_profit[i] + annual_dep - principal[i]

    def calc_irr(cf):
        try:
            r = 0.1
            for _ in range(1000):
                npv = sum(c / (1 + r) ** t for t, c in enumerate(cf))
                dnpv = sum(-t * c / (1 + r) ** (t + 1) for t, c in enumerate(cf))
                if abs(dnpv) < 1e-12:
                    break
                r_new = r - npv / dnpv
                if abs(r_new - r) < 1e-8:
                    r = r_new
                    break
                r = r_new
            return r * 100
        except:
            return None

    full_irr = calc_irr(fcf)
    equity_irr = calc_irr(eq_cf) if leverage < 1.0 else None

    # LCOE = (CapEx + Σ(OpEx_t+Tax_t)/(1+WACC)^t) / Σ(Gen_t/(1+WACC)^t)
    total_cost_pv = ti
    total_gen_pv = 0
    for i in range(operating_years):
        yr_cost = om[i] + ins[i] + income_tax[i] + add_tax[i]
        total_cost_pv += yr_cost / (1 + wacc) ** (i + 1)
        total_gen_pv += gen[i] / (1 + wacc) ** (i + 1)
    lcoe = total_cost_pv / total_gen_pv if total_gen_pv > 0 else 0

    min_dscr = min(dscr[:loan_years])

    avail_for_ds = ebitda - add_tax - income_tax

    return {
        'unit_inv': unit_inv,
        'ti': ti,
        'debt': debt,
        'equity': equity,
        'min_dscr': min_dscr,
        'full_irr': full_irr,
        'equity_irr': equity_irr,
        'lcoe': lcoe,
        'avg_ncf': np.mean(net_cf),
        'avg_np': np.mean(net_profit),
        'annual_dep': annual_dep,
        'annual_principal_amt': annual_principal_amt,
        # 年度明细数组
        'gen': gen,
        'rev': rev,
        'om': om,
        'ins': ins,
        'ebitda': ebitda,
        'interest': interest,
        'principal': principal,
        'add_tax': add_tax,
        'income_tax': income_tax,
        'net_profit': net_profit,
        'dscr': dscr,
        'net_cf': net_cf,
        'avail_for_ds': avail_for_ds,
        'nbv': nbv_arr,
    }


def bin_search_dscr(target, leverage, lo, hi):
    """二分搜索：最小DSCR ≥ target 下的最高单瓦投资"""
    best = lo
    for _ in range(60):
        mid = (lo + hi) / 2
        r = calculate(mid, leverage)
        if r['min_dscr'] >= target:
            best = mid
            lo = mid
        else:
            hi = mid
        if hi - lo < 0.0001:
            break
    return round(best, 4)


def bin_search_dual(full_irr_tgt, eq_irr_tgt, lo, hi):
    """二分搜索：全投资IRR≥目标 且 资本金IRR≥目标 下的最高单瓦投资（80%融资）"""
    best = lo
    for _ in range(60):
        mid = (lo + hi) / 2
        r = calculate(mid, 0.8)
        ok = r['full_irr'] is not None and r['full_irr'] >= full_irr_tgt
        ok = ok and r['equity_irr'] is not None and r['equity_irr'] >= eq_irr_tgt
        if ok:
            best = mid
            lo = mid
        else:
            hi = mid
        if hi - lo < 0.0001:
            break
    return round(best, 4)


def print_profit_table(r, title, leverage):
    """打印每年净利润表"""
    print(f"\n{'─'*130}")
    print(f"  {title} — 每年净利润表（单位：万元）")
    print(f"{'─'*130}")
    header = (f"{'年份':>6} {'发电量(MWh)':>12} {'营业收入':>10} {'运维费':>8} "
              f"{'保险费':>8} {'折旧':>8} {'利息':>8} {'增值税及附加':>10} "
              f"{'所得税':>8} {'净利润':>8}")
    print(header)
    print(f"{'─'*130}")
    for i in range(operating_years):
        y = i + 1
        print(f"{y:>6} "
              f"{r['gen'][i]/1000:>12.1f} "
              f"{r['rev'][i]/1e4:>10.2f} "
              f"{r['om'][i]/1e4:>8.2f} "
              f"{r['ins'][i]/1e4:>8.2f} "
              f"{r['annual_dep']/1e4:>8.2f} "
              f"{r['interest'][i]/1e4:>8.2f} "
              f"{r['add_tax'][i]/1e4:>10.2f} "
              f"{r['income_tax'][i]/1e4:>8.2f} "
              f"{r['net_profit'][i]/1e4:>8.2f}")
    print(f"{'─'*130}")
    print(f"  公式：营业收入=年发电量×有效电价 | 发电量含逐年衰减({degradation_rate*100:.1f}%/年)")
    print(f"  运维费=装机×单位费率(分段) + 第10年逆变器替换 | 保险费=净值×0.2%")
    print(f"  折旧=总投资×(1-残值率{residual_rate*100:.0f}%)/{operating_years}年 | 利息=期初余额×{interest_rate*100:.0f}%")
    print(f"  增值税=销项税 - 进项税(融资租赁租金本息×{vat_rate*100:.0f}% + 运维费×{om_vat_rate*100:.0f}%)")
    print(f"  所得税=max(0,(收入-运维-保险-折旧-利息-增值税及附加)×{income_tax_rate*100:.0f}%)")
    print(f"  净利润=营业收入-运维费-保险费-折旧-利息-增值税及附加-所得税")


def print_cf_table(r, title, leverage):
    """打印每年净现金流量表"""
    print(f"\n{'─'*130}")
    print(f"  {title} — 每年净现金流量表（单位：万元）")
    print(f"{'─'*130}")
    if leverage >= 1.0:
        header = (f"{'年份':>6} {'EBITDA':>10} {'增值税及附加':>10} {'所得税':>8} "
                  f"{'可用于还款':>10} {'应还本金':>10} {'应还利息':>10} "
                  f"{'应还本息':>10} {'偿债备付率':>8} {'净现金流':>10}")
        print(header)
    else:
        header = (f"{'年份':>6} {'EBITDA':>10} {'增值税及附加':>10} {'所得税':>8} "
                  f"{'可用于还款':>10} {'应还本金':>10} {'应还利息':>10} "
                  f"{'应还本息':>10} {'偿债备付率':>8} {'股权现金流':>10}")
        print(header)
    print(f"{'─'*130}")
    for i in range(operating_years):
        y = i + 1
        ds = r['principal'][i] + r['interest'][i]
        dscr_val = f"{r['dscr'][i]:>8.2f}" if r['dscr'][i] != float('inf') else f"{'∞':>8}"
        if leverage >= 1.0:
            cf_val = r['net_cf'][i]
        else:
            cf_val = r['net_profit'][i] + r['annual_dep'] - r['principal'][i]
        print(f"{y:>6} "
              f"{r['ebitda'][i]/1e4:>10.2f} "
              f"{r['add_tax'][i]/1e4:>10.2f} "
              f"{r['income_tax'][i]/1e4:>8.2f} "
              f"{r['avail_for_ds'][i]/1e4:>10.2f} "
              f"{r['principal'][i]/1e4:>10.2f} "
              f"{r['interest'][i]/1e4:>10.2f} "
              f"{ds/1e4:>10.2f} "
              f"{dscr_val} "
              f"{cf_val/1e4:>10.2f}")
    print(f"{'─'*130}")
    print(f"  公式：EBITDA=营业收入-运维费-保险费 | 可用于还款=EBITDA-增值税及附加-所得税")
    print(f"  当年应还本=期初本金/{loan_years}年 | 应还利息=期初余额×{interest_rate*100:.0f}%")
    print(f"  偿债备付率=可用于还款/应还本息")
    if leverage >= 1.0:
        print(f"  净现金流量=EBITDA-增值税及附加-所得税-当年利息-当年偿还本金")
    else:
        print(f"  股权现金流=净利润+折旧-当年偿还本金 | 全投资IRR使用无杠杆附加税计算")


if __name__ == '__main__':
    print(f"{'='*80}")
    print(f"光伏项目投资评估 —— 投资边界 & 出售边界")
    print(f"{'='*80}")
    print(f"装机容量: {capacity_mw}MW（直流侧）")
    print(f"等效利用小时数: {annual_hours}h")
    print(f"限电率: {curtailment_rate*100:.2f}%")
    print(f"组件衰减率: {degradation_rate*100:.2f}%/年")
    base_gen_mwh = base_annual_gen_kwh / 1000
    print(f"首年净发电量: {base_gen_mwh:,.0f}MWh")
    print(f"有效电价: {effective_price_base:.4f}元/kWh")
    print(f"融资利率: {interest_rate*100:.0f}% | 期限: {loan_years}年 | 经营期: {operating_years}年")
    print(f"迭代范围: {search_lo}~{search_hi}元/W")

    # 任务1：投资边界
    print(f"\n{'='*80}")
    print(f"  任务1：投资边界（100%融资，最小DSCR≥1.2）")
    print(f"{'='*80}")

    t1 = bin_search_dscr(1.2, 1.0, search_lo, search_hi)
    r1 = calculate(t1, 1.0)

    print(f"\n  ▶ 最高单瓦投资: {r1['unit_inv']:.4f}元/W")
    print(f"  ▶ 总投资: {r1['ti']/1e8:.4f}亿元")
    print(f"  ▶ 最小DSCR: {r1['min_dscr']:.4f}")
    print(f"  ▶ 全投资IRR: {r1['full_irr']:.2f}%")
    print(f"  ▶ LCOE: {r1['lcoe']:.4f}元/kWh")
    print(f"  ▶ 年均净现金流量: {r1['avg_ncf']/1e4:.2f}万元")
    print(f"  ▶ 年均净利润: {r1['avg_np']/1e4:.2f}万元")

    print_profit_table(r1, "任务1·投资边界", 1.0)
    print_cf_table(r1, "任务1·投资边界", 1.0)

    # 任务4：出售边界
    print(f"\n{'='*80}")
    print(f"  任务4：出售边界（80%融资，全投资IRR≥6% 且 资本金IRR≥8%）")
    print(f"{'='*80}")

    t4 = bin_search_dual(6.0, 8.0, search_lo, search_hi)
    r4 = calculate(t4, 0.8)

    print(f"\n  ▶ 目标单瓦投资: {r4['unit_inv']:.4f}元/W")
    print(f"  ▶ 总投资: {r4['ti']/1e8:.4f}亿元")
    print(f"  ▶ 全投资IRR: {r4['full_irr']:.2f}%")
    print(f"  ▶ 税后资本金IRR: {r4['equity_irr']:.2f}%")
    print(f"  ▶ LCOE: {r4['lcoe']:.4f}元/kWh")
    print(f"  ▶ 年均净现金流量: {r4['avg_ncf']/1e4:.2f}万元")
    print(f"  ▶ 年均净利润: {r4['avg_np']/1e4:.2f}万元")

    print_profit_table(r4, "任务4·出售边界", 0.8)
    print_cf_table(r4, "任务4·出售边界", 0.8)

    # 对比总结
    print(f"\n{'='*80}")
    print(f"  投资边界 vs 出售边界 对比")
    print(f"{'='*80}")
    print(f"  {'指标':<24} {'投资边界':>14} {'出售边界':>14}")
    print(f"  {'─'*52}")
    print(f"  {'融资结构':<24} {'100%融资':>14} {'80%融资+20%资本金':>14}")
    print(f"  {'约束条件':<24} {'DSCR≥1.2':>14} {'全投资IRR≥6% +':>14}")
    print(f"  {'':<24} {'':>14} {'资本金IRR≥8%':>14}")
    print(f"  {'单瓦投资(元/W)':<24} {r1['unit_inv']:>14.4f} {r4['unit_inv']:>14.4f}")
    print(f"  {'总投资(亿元)':<24} {r1['ti']/1e8:>14.4f} {r4['ti']/1e8:>14.4f}")
    print(f"  {'全投资IRR':<24} {r1['full_irr']:>13.2f}% {r4['full_irr']:>13.2f}%")
    print(f"  {'资本金IRR':<24} {'─':>14} {r4['equity_irr']:>13.2f}%")
    print(f"  {'最小DSCR':<24} {r1['min_dscr']:>14.4f} {r4['min_dscr']:>14.4f}")
    print(f"  {'LCOE(元/kWh)':<24} {r1['lcoe']:>14.4f} {r4['lcoe']:>14.4f}")
    print(f"  {'年均净利润(万元)':<24} {r1['avg_np']/1e4:>14.2f} {r4['avg_np']/1e4:>14.2f}")
