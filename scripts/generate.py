#!/usr/bin/env python3
"""
generate.py
讀取 data/ 資料夾內所有銷售檔案，產出 index.html
執行方式: python scripts/generate.py
"""

import os
import re
import json
import glob
from datetime import date
import pandas as pd

# ── 路徑設定 ──────────────────────────────────────────
ROOT        = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR    = os.path.join(ROOT, 'data')
STORE_FILE  = os.path.join(DATA_DIR, '門市基本資料.xlsx')
OUTPUT_FILE = os.path.join(ROOT, 'index.html')
SCHED_FILE  = os.path.join(DATA_DIR, 'AIRSPACE_線下通路_行銷排程_含甘特圖_2025-2026.xlsx')

# ── 折扣欄位 ──────────────────────────────────────────
DISCOUNT_COLS = [
    '會員折扣攤算','滿額折現攤算','折扣碼攤算','折價券攤算',
    '區間折扣金額攤算','任選折扣金額攤算','紅利金攤算','會員生日年度折扣攤算'
]
DISC_TYPE_MAP = {
    '任選折扣金額攤算': '任選折(複數)',
    '區間折扣金額攤算': '區間折(單件)',
    '紅利金攤算':       '紅利折抵',
    '滿額折現攤算':     '滿額折現',
    '折扣碼攤算':       '折扣碼',
    '會員折扣攤算':     '會員折扣',
    '折價券攤算':       '折價券',
}
LEVEL_MAP = {'00':'過路客','01':'一般會員','02':'VIP','03':'SVIP','09':'員購'}
MEMBER_THRESHOLD = {'00':8,'01':8,'02':11,'03':18}
PROD_ORDER = ['正品','單一特','組合價','續賣品','出清品','特惠品']
PROD_THRESHOLD = {'正品':10,'單一特':28,'組合價':21,'續賣品':10}
EXCL_PROD = {'出清品','特惠品'}

# ── 讀門市基本資料 ────────────────────────────────────
def load_store():
    df = pd.read_excel(STORE_FILE, sheet_name='店數')
    df = df[['倉庫名稱','販售支線','門市型態']].dropna(subset=['販售支線'])
    # AS62 補街邊店
    df.loc[df['倉庫名稱'].str.contains('南西IP|南西三', na=False), '門市型態'] = '街邊店'
    return df

# ── 讀行銷排程 ────────────────────────────────────────
def load_sched(year: int, month: int):
    if not os.path.exists(SCHED_FILE):
        return []
    df = pd.read_excel(SCHED_FILE, sheet_name='行銷排程')
    df = df[
        (df['年度'] == year) &
        (df['月份'] == f'{month}月') &
        (df['狀態'] != '取消')          # 過濾取消活動
    ][['起訖日期','活動類型','活動內容','適用門市/通路']].copy()
    df = df.fillna('')
    df = df.rename(columns={'適用門市/通路':'適用門市'})
    # 去重（排程有時會有重複列）
    df = df.drop_duplicates(subset=['起訖日期','活動內容'])
    # 起訖日期可能是 datetime 格式，統一轉字串
    df['起訖日期'] = df['起訖日期'].apply(
        lambda x: x.strftime('%m/%d') if hasattr(x, 'strftime') else str(x)
    )
    return df.to_dict('records')

# ── 處理單一月份銷售資料 ─────────────────────────────
def process_month(filepath: str, store_df: pd.DataFrame) -> dict:
    df = pd.read_excel(filepath)

    # 排除免費贈品（小計=0 且 零售價=0）
    df = df[~((df['小計'] == 0) & (df['零售價'] == 0))].copy()

    # 折扣合計
    df['折扣合計'] = df[DISCOUNT_COLS].apply(lambda x: abs(x)).sum(axis=1)

    # 會員等級碼
    df['會員等級碼'] = df['購買當下會員等級'].astype(str).str.zfill(2)

    # 合併門市資料
    df = df.merge(store_df, on='倉庫名稱', how='left')
    df['販售支線'] = df['販售支線'].fillna('其他')
    df['門市型態'] = df['門市型態'].fillna('街邊店')
    df['通路類型'] = df['倉庫名稱'].apply(
        lambda x: 'OUTLET' if 'OUTLET' in str(x) else '正價'
    )

    # 行促類型調整：空白有折扣 → 會員行促
    df['行促類型_adj'] = df['行促類型'].fillna('') if '行促類型' in df.columns else ''
    df.loc[(df['行促類型_adj'] == '') & (df['折扣合計'] > 0), '行促類型_adj'] = '會員行促'
    df.loc[(df['行促類型_adj'] == '') & (df['折扣合計'] == 0), '行促類型_adj'] = '無折扣'

    # 分群
    df_main   = df[(df['通路類型'] == '正價') & (df['會員等級碼'] != '09')].copy()
    df_outlet = df[df['通路類型'] == 'OUTLET'].copy()
    df_emp    = df[df['會員等級碼'] == '09'].copy()
    df_nooutlet_noemp = df[df['會員等級碼'] != '09'].copy()  # 含正價+OUTLET，排員購

    # ── KPI ──
    sell  = df_main['零售價'].sum()
    rev   = df_main['小計'].sum()
    disc  = df_main['折扣合計'].sum()
    orders       = df_main['訂單編號'].nunique()
    disc_orders  = df_main[df_main['折扣合計'] > 0]['訂單編號'].nunique()
    ord_agg = df_main.groupby('訂單編號').agg(
        有折扣=('折扣合計', lambda x: (x > 0).any()),
        客單=('小計', 'sum')
    ).reset_index()
    avg_d = ord_agg[ord_agg['有折扣'] == True]['客單'].mean()
    avg_n = ord_agg[ord_agg['有折扣'] == False]['客單'].mean()

    sell_all = df_nooutlet_noemp['零售價'].sum() + df_outlet['零售價'].sum()
    rev_all  = df_nooutlet_noemp['小計'].sum()   + df_outlet['小計'].sum()
    disc_all = df_nooutlet_noemp['折扣合計'].sum() + df_outlet['折扣合計'].sum()

    kpi = {
        'disc_rate':      round(disc / sell * 100, 1) if sell else 0,
        'disc_all_rate':  round(disc_all / sell_all * 100, 1) if sell_all else 0,
        'disc_order_pct': round(disc_orders / orders * 100, 1) if orders else 0,
        'avg_disc':       round(avg_d) if pd.notna(avg_d) else 0,
        'avg_nodis':      round(avg_n) if pd.notna(avg_n) else 0,
        'uplift':         round((avg_d / avg_n - 1) * 100, 1) if pd.notna(avg_d) and pd.notna(avg_n) and avg_n else 0,
        'sell':           int(sell),
        'rev':            int(rev),
        'sell_all':       int(sell_all),
        'rev_all':        int(rev_all),
    }

    # ── 折扣類型 ──
    total_disc = df_main['折扣合計'].sum()
    disc_types = []
    for col, lbl in DISC_TYPE_MAP.items():
        amt = abs(df_main[col]).sum()
        if amt > 0:
            disc_types.append({
                'name': lbl,
                'amt':  int(amt),
                'pct':  round(amt / total_disc * 100, 1) if total_disc else 0,
            })
    disc_types.sort(key=lambda x: -x['amt'])

    # ── 品別折扣率 ──
    prod_grp = df_main.groupby('品別').agg(
        sell=('零售價', 'sum'), disc=('折扣合計', 'sum')
    ).reset_index()
    prod_grp['disc_rate'] = (prod_grp['disc'] / prod_grp['sell'] * 100).round(1)
    prod_dict = {r['品別']: r for _, r in prod_grp.iterrows()}
    prod_type = []
    for p in PROD_ORDER:
        if p in prod_dict:
            r = prod_dict[p]
            prod_type.append({
                '品別':      p,
                'sell':      int(r['sell']),
                'disc_rate': float(r['disc_rate']),
                'excl':      p in EXCL_PROD,
                'threshold': PROD_THRESHOLD.get(p, None),
            })

    # ── 會員等級 ──
    members = []
    for code in ['00', '01', '02', '03']:
        sub = df_main[df_main['會員等級碼'] == code]
        if len(sub) == 0:
            continue
        s = sub['零售價'].sum()
        d = sub['折扣合計'].sum()
        r = sub['小計'].sum()
        bk = {}
        for col, lbl in DISC_TYPE_MAP.items():
            amt = abs(sub[col]).sum()
            if amt > 0:
                bk[lbl] = round(amt / s * 100, 1)
        members.append({
            'code':      code,
            'name':      LEVEL_MAP[code],
            'sell':      int(s),
            'rev':       int(r),
            'disc_rate': round(d / s * 100, 1) if s else 0,
            'threshold': MEMBER_THRESHOLD[code],
            'breakdown': bk,
        })

    # ── 行促類型 ──
    pt = df_main[df_main['折扣合計'] > 0].groupby('行促類型_adj').agg(
        sell=('零售價', 'sum'), disc=('折扣合計', 'sum'),
        rev=('小計', 'sum'), 筆數=('訂單編號', 'count')
    ).reset_index()
    pt['disc_rate'] = (pt['disc'] / pt['sell'] * 100).round(1)
    pt['pct']       = (pt['disc'] / total_disc * 100).round(1) if total_disc else 0
    pt = pt.sort_values('disc', ascending=False)
    promo_type = pt.rename(columns={'行促類型_adj': 'name'})[
        ['name', 'sell', 'rev', 'disc', 'disc_rate', 'pct']
    ].to_dict('records')
    for p in promo_type:
        p['sell'] = int(p['sell']); p['rev'] = int(p['rev']); p['disc'] = int(p['disc'])

    # ── 行促名稱（前15，有折扣金額）──
    if '行促名稱' in df_main.columns:
        pn = df_main[df_main['折扣合計'] > 0].groupby('行促名稱').agg(
            sell=('零售價', 'sum'), disc=('折扣合計', 'sum'),
            rev=('小計', 'sum'), cnt=('訂單編號', 'count')
        ).reset_index()
        pn = pn[pn['行促名稱'].notna() & (pn['行促名稱'] != '')].copy()
        pn['disc_rate'] = (pn['disc'] / pn['sell'] * 100).round(1)
        pn['pct']       = (pn['disc'] / total_disc * 100).round(1) if total_disc else 0
        pn = pn.sort_values('disc', ascending=False).head(15)
        promo_name = pn.rename(columns={'行促名稱': 'name'})[
            ['name', 'sell', 'rev', 'cnt', 'disc_rate', 'pct']
        ].to_dict('records')
        for p in promo_name:
            p['sell'] = int(p['sell']); p['rev'] = int(p['rev']); p['cnt'] = int(p['cnt'])
    else:
        promo_name = []

    # ── 門市型態（排員購）──
    st = df_main.groupby('門市型態').agg(
        sell=('零售價', 'sum'), disc=('折扣合計', 'sum'), rev=('小計', 'sum')
    ).reset_index()
    st['disc_rate'] = (st['disc'] / st['sell'] * 100).round(1)
    store_type = st.rename(columns={'門市型態': 'name'})[
        ['name', 'sell', 'rev', 'disc_rate']
    ].to_dict('records')
    for s in store_type:
        s['sell'] = int(s['sell']); s['rev'] = int(s['rev'])

    # ── 販售支線（排員購）──
    br = df_main.groupby('販售支線').agg(
        sell=('零售價', 'sum'), disc=('折扣合計', 'sum'), rev=('小計', 'sum')
    ).reset_index()
    br['disc_rate'] = (br['disc'] / br['sell'] * 100).round(1)
    br = br.sort_values('disc_rate', ascending=False)
    branch = br.rename(columns={'販售支線': 'name'})[
        ['name', 'sell', 'rev', 'disc_rate']
    ].to_dict('records')
    for b in branch:
        b['sell'] = int(b['sell']); b['rev'] = int(b['rev'])

    # ── 參考區 ──
    o_sell = int(df_outlet['零售價'].sum())
    o_rev  = int(df_outlet['小計'].sum())
    o_disc = df_outlet['折扣合計'].sum()
    e_sell = int(df_emp['零售價'].sum())
    e_rev  = int(df_emp['小計'].sum())
    e_disc = df_emp['折扣合計'].sum()

    ref = {
        'outlet': {
            'sell': o_sell, 'rev': o_rev,
            'disc_rate': round(o_disc / o_sell * 100, 1) if o_sell else 0
        },
        'emp': {
            'sell': e_sell, 'rev': e_rev,
            'disc_rate': round(e_disc / e_sell * 100, 1) if e_sell else 0
        }
    }

    # ── 門市警示：正價前5高折扣率 ──
    store_grp = df_main.groupby('倉庫名稱').agg(
        sell=('零售價', 'sum'), disc=('折扣合計', 'sum')
    ).reset_index()
    store_grp['disc_rate'] = (store_grp['disc'] / store_grp['sell'] * 100).round(1)
    store_grp = store_grp[store_grp['sell'] >= 300000]
    avg_store_rate = round(store_grp['disc_rate'].mean(), 1)
    top5_stores = store_grp.nlargest(5, 'disc_rate')[['倉庫名稱', 'disc_rate']].reset_index(drop=True)
    store_warning = []
    for _, row in top5_stores.iterrows():
        vs = round(row['disc_rate'] - avg_store_rate, 1)
        store_warning.append({
            'name':      row['倉庫名稱'],
            'disc_rate': float(row['disc_rate']),
            'vs_avg':    vs,
        })

    # ── 當月建議觸發警示 ──
    svip = next((m for m in members if m['code'] == '03'), None)
    svip_bonus = svip['breakdown'].get('紅利折抵', 0) if svip else 0
    任選pct = next((t['pct'] for t in disc_types if t['name'] == '任選折(複數)'), 0)
    正品rate = next((p['disc_rate'] for p in prod_type if p['品別'] == '正品'), 0)

    alerts = []
    if svip_bonus > 8:
        alerts.append({'type': 'svip_bonus', 'val': svip_bonus})
    if 任選pct > 50:
        alerts.append({'type': 'promo_pct', 'val': 任選pct})
    if 正品rate > 10:
        alerts.append({'type': 'normal_rate', 'val': 正品rate})
    if kpi['disc_order_pct'] > 65:
        alerts.append({'type': 'order_pct', 'val': kpi['disc_order_pct']})
    branch_avg = sum(b['disc_rate'] for b in branch) / len(branch) if branch else 0
    for b in branch:
        if b['disc_rate'] - branch_avg >= 2:
            alerts.append({'type': 'branch_high', 'name': b['name'], 'val': b['disc_rate'], 'avg': round(branch_avg, 1)})

    return {
        'kpi':           kpi,
        'discTypes':     disc_types,
        'members':       members,
        'prodType':      prod_type,
        'promoType':     promo_type,
        'promoName':     promo_name,
        'storeType':     store_type,
        'branch':        branch,
        'ref':           ref,
        'storeWarning':  store_warning,
        'storeAvgRate':  avg_store_rate,
        'alerts':        alerts,
        'recs':          [],
        'sched':         [],
    }

# ── 自動建議產生（當月 + 3個月趨勢）────────────────
def generate_recs(all_data: dict) -> dict:
    keys = sorted(all_data.keys())
    recs_by_month = {}

    for i, key in enumerate(keys):
        d = all_data[key]
        alerts = d.get('alerts', [])
        # 最近3個月 keys（含本月）
        recent_keys = keys[max(0, i-2): i+1]
        recs = []

        # ① SVIP 紅利折抵
        svip_bonus_vals = []
        for k in recent_keys:
            svip = next((m for m in all_data[k]['members'] if m['code']=='03'), None)
            if svip:
                svip_bonus_vals.append(svip['breakdown'].get('紅利折抵', 0))
        if svip_bonus_vals and svip_bonus_vals[-1] > 8:
            months_high = sum(1 for v in svip_bonus_vals if v > 8)
            trend = f"連續 {months_high} 個月 > 8%" if months_high > 1 else f"本月 {svip_bonus_vals[-1]}%"
            recs.append({
                'title': '① 紅利點數管理',
                'sub':   f'SVIP 紅利折抵偏高（{trend}）',
                'items': [
                    f'SVIP 紅利折抵本月達 {svip_bonus_vals[-1]}%，建議設定每筆最高折抵上限（如訂單 10%）',
                    '縮短點數有效期，降低累積後單次大額折抵',
                    '高級距點數改換贈品或體驗，保住毛利同時維持 SVIP 黏著度',
                ]
            })

        # ② 任選折佔比
        promo_vals = []
        for k in recent_keys:
            pct = next((t['pct'] for t in all_data[k]['discTypes'] if t['name']=='任選折(複數)'), 0)
            promo_vals.append(pct)
        if promo_vals and promo_vals[-1] > 50:
            months_high = sum(1 for v in promo_vals if v > 50)
            trend = f"連續 {months_high} 個月 > 50%" if months_high > 1 else f"本月 {promo_vals[-1]}%"
            recs.append({
                'title': '② 活動結構調整',
                'sub':   f'任選折佔總折扣偏高（{trend}）',
                'items': [
                    '每月設至少兩週活動空窗，讓消費者回到原價體驗',
                    '建立新品保護期（上市前 4 週不列入任選折）',
                    '以滿額贈替代折現活動，維持客單不傷折扣率',
                ]
            })

        # ③ 正品折扣率
        normal_vals = []
        for k in recent_keys:
            r = next((p['disc_rate'] for p in all_data[k]['prodType'] if p['品別']=='正品'), 0)
            normal_vals.append(r)
        if normal_vals and normal_vals[-1] > 10:
            months_high = sum(1 for v in normal_vals if v > 10)
            trend = f"連續 {months_high} 個月 > 10%" if months_high > 1 else f"本月 {normal_vals[-1]}%"
            recs.append({
                'title': '③ 正品折扣率管控',
                'sub':   f'正品折扣率超標（{trend}）',
                'items': [
                    f'正品折扣率本月 {normal_vals[-1]}%，建議以 10% 為上限',
                    '出清品與正品分開追蹤，避免稀釋正品指標',
                    '員購折扣隔離計算，不納入對外報告',
                ]
            })

        # ④ 支線連續偏高（3個月趨勢）
        if len(recent_keys) >= 2:
            # 找各月各支線折扣率
            branch_history = {}
            for k in recent_keys:
                for b in all_data[k]['branch']:
                    if b['name'] not in branch_history:
                        branch_history[b['name']] = []
                    branch_history[b['name']].append(b['disc_rate'])
            # 計算各月整體均值
            overall_avgs = []
            for k in recent_keys:
                rates = [b['disc_rate'] for b in all_data[k]['branch']]
                overall_avgs.append(sum(rates)/len(rates) if rates else 0)
            # 找連續偏高支線
            high_branches = []
            for bname, brates in branch_history.items():
                if len(brates) == len(recent_keys):
                    above = [r - overall_avgs[j] >= 2 for j, r in enumerate(brates)]
                    if sum(above) >= 2:
                        high_branches.append({'name': bname, 'rates': brates, 'months': sum(above)})
            if high_branches:
                branch_list = '、'.join([f"{b['name']}（{b['rates'][-1]}%）" for b in high_branches[:3]])
                recs.append({
                    'title': '④ 支線折扣率持續偏高',
                    'sub':   f'以下支線近期持續高於整體均值 +2pp 以上',
                    'items': [
                        f'{branch_list} 折扣率偏高，建議確認是否有活動重疊',
                        '檢視該支線適用活動是否過多，設定互斥規則',
                        '若無法調整活動，建議提高該支線正常品售價作為緩衝',
                    ]
                })

        # ⑤ 有折扣訂單佔比
        if d['kpi']['disc_order_pct'] > 65:
            recs.append({
                'title': '⑤ 折扣覆蓋率過高',
                'sub':   f"本月 {d['kpi']['disc_order_pct']}% 訂單享有折扣",
                'items': [
                    '超過六成訂單享折扣，消費者容易養成等折扣習慣',
                    '建議縮短活動天數，增加原價購買機會',
                    '搭配滿額贈取代直接折現，降低折扣依賴',
                ]
            })

        # 若無觸發，給預設健康提示
        if not recs:
            recs.append({
                'title': '✅ 本月折扣結構健康',
                'sub':   '各項指標均在門檻值內',
                'items': [
                    '持續維持正品折扣率在 10% 以下',
                    '留意 SVIP 紅利折抵趨勢，避免累積後單次大額折抵',
                    '活動空窗期維持，保持消費者原價購買習慣',
                ]
            })

        recs_by_month[key] = recs

    return recs_by_month


# ── 掃描 data/ 資料夾 ────────────────────────────────
def scan_files():
    pattern = os.path.join(DATA_DIR, '銷售_????_?月.xlsx')
    files = glob.glob(pattern)
    # 也抓兩位數月份：銷售_2026_10月.xlsx
    pattern2 = os.path.join(DATA_DIR, '銷售_????_??月.xlsx')
    files += glob.glob(pattern2)
    results = []
    for f in sorted(set(files)):
        m = re.search(r'銷售_(\d{4})_(\d{1,2})月\.xlsx', os.path.basename(f))
        if m:
            results.append((int(m.group(1)), int(m.group(2)), f))
    results.sort()
    return results

# ── 讀 HTML 模板（v9 的前後兩段）────────────────────
def load_template():
    tmpl_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'template.html')
    return open(tmpl_path, encoding='utf-8').read()

# ── 主程式 ───────────────────────────────────────────
def main():
    print("📂 掃描 data/ 資料夾...")
    files = scan_files()
    if not files:
        print("❌ 找不到銷售資料，請確認 data/ 資料夾內有 銷售_YYYY_M月.xlsx 檔案")
        return

    print(f"✅ 找到 {len(files)} 個月份：{[f'{y}-{m:02d}' for y,m,_ in files]}")

    print("📋 讀取門市基本資料...")
    store_df = load_store()

    all_data = {}
    for year, month, filepath in files:
        key = f"{year}-{month:02d}"
        print(f"⚙️  處理 {key} ({os.path.basename(filepath)})...")
        try:
            data = process_month(filepath, store_df)
            data['sched'] = load_sched(year, month)
            all_data[key] = data
        except Exception as e:
            print(f"⚠️  {key} 處理失敗：{e}")

    # 計算趨勢建議
    print("💡 產生自動建議...")
    recs_by_month = generate_recs(all_data)
    for key in all_data:
        all_data[key]['recs'] = recs_by_month.get(key, [])

    print("📝 產出 index.html...")
    template = load_template()

    # 把 DATA = {} 注入 template
    data_json = json.dumps(all_data, ensure_ascii=False, indent=None)
    today = date.today().strftime('%Y-%m-%d')
    html = template.replace('{{DATA}}', data_json)
    html = html.replace('{{UPDATED}}', today)

    open(OUTPUT_FILE, 'w', encoding='utf-8').write(html)
    print(f"✅ 完成！輸出至 {OUTPUT_FILE}")
    print(f"   月份：{list(all_data.keys())}")

if __name__ == '__main__':
    main()
